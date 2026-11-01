#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# """LOGIDATA BOGOTÁ — FASES I Y II INTEGRADAS (único archivo).
#
# Dependencias: pip install pandas numpy scipy openpyxl matplotlib
# Ejemplo: python LogiData_Fases_I_II_integradas.py --output LogiData_consolidado.xlsx
#
# Fase I: construcción híbrida original, múltiples semillas, cierres y 72
# construcciones POR (semilla, cierre). Retiene el TOP 5 estructural exacto,
# sin duplicados, entre todas las soluciones factibles de TODAS las construcciones.
# Fase II: 2-opt, Relocate, Swap, fusión y reasignación; mejor mejora por barrido.
# No construye nuevamente una solución de Fase I, no cambia las bodegas abiertas.
# Ambas fases se transfieren en RAM y se exportan en UN libro Excel.
# """
from __future__ import annotations
# ======================== MÓDULO CONSERVADO: config.py ========================
# """Parámetros de carga y ejecución modificables en este archivo único."""
from dataclasses import dataclass, field
import math


@dataclass
class InputConfig:
    depots_file: str
    clients_file: str
    distance_file: str
    time_file: str
    depots_sheet: str = 'Bodegas'
    clients_sheet: str = 'Clientes'
    distance_sheet: str = 'Distancias'
    time_sheet: str = 'Tiempos'
    depot_id: str = 'id'
    depot_capacity: str = 'capacidad'
    depot_cost: str = 'costo_apertura'
    # 'investment': columna original en COP de inversión; 'daily': ya está en COP/día.
    depot_cost_basis: str = 'investment'
    depot_useful_life_years: float = 15.0
    days_per_year: int = 365
    client_id: str = 'id'
    delivery: str = 'entrega'
    pickup: str = 'recogida'
    # El encabezado de la primera columna de la matriz (etiquetas de filas)
    matrix_row_id: str = 'id'
    # Columnas opcionales; si no existen se usa visualización esquemática.
    depot_x: str = 'x'
    depot_y: str = 'y'
    client_x: str = 'x'
    client_y: str = 'y'
    # Los valores de tiempos se interpretan en minutos.


@dataclass
class SearchConfig:
    vehicle_capacity: float = 2500.0
    cost_per_km: float = 5000.0
    vehicle_rental_daily: float = 147572.55  # Renault Trafic L2H2: COP/vehículo/día
    # Solo DOE/ejecución fija: número de CONFIGURACIONES a ensayar.
    combinations: int = 50
    # Solo DOE/ejecución fija: arranques por configuración seleccionada.
    starts_per_combination: int = 10
    execution_mode: str = 'budget'  # 'budget': producción reproducible; 'fixed': DOE
    construction_budget: int = 72  # seleccionado con 10 semillas: 100 no demostró ahorro medio significativo
    time_limit_seconds: float = 10.0  # solo legado de estudios históricos; NO detiene la búsqueda
    geometric_p: float = 0.95  # probabilidad del mejor en distribución geométrica truncada
    seed: int = 4110
    # Compatibilidad con configuración previa; no se usan al enumerar TODO.
    enumeration_threshold: int = 50000
    random_draw_limit: int = 20000
    assignment_repairs: int = 2
    cluster_split_ratio: float = 0.65  # reserva si no existe calibración validada para estos Excel
    keep_top: int = 10
    # Exigir distintos conjuntos de bodegas o al menos este porcentaje de clientes
    # reasignados/rutas modificadas si hay suficientes alternativas.
    diversity_fraction: float = 0.10
    diversity_cost_tolerance: float = 0.20
    output_excel: str = 'logidata_fase1_resultados.xlsx'
    output_png: str = 'logidata_mejor_solucion.png'

    def validate(self) -> None:
        if self.vehicle_capacity <= 0 or self.cost_per_km < 0 or self.vehicle_rental_daily < 0:
            raise ValueError('Capacidad vehicular debe ser positiva y costo/km no negativo')
        if self.combinations < 1 or self.starts_per_combination < 1 or self.keep_top < 1:
            raise ValueError('combinations, starts_per_combination y keep_top deben ser >= 1')
        if self.execution_mode not in ('budget', 'fixed'):
            raise ValueError('execution_mode debe ser budget o fixed')
        if isinstance(self.construction_budget, bool) or not isinstance(self.construction_budget, int) or self.construction_budget < 1:
            raise ValueError('construction_budget debe ser un entero >= 1')
        # time_limit_seconds queda solo como dato de compatibilidad histórica.
        # Nunca se consulta durante la búsqueda por construcciones.
        if not (0 < self.geometric_p <= 1):
            raise ValueError('geometric_p debe pertenecer a (0,1]')
        if self.enumeration_threshold < 1 or self.random_draw_limit < 1:
            raise ValueError('Límites de exploración deben ser positivos')
        if not 0 < self.cluster_split_ratio <= 1:
            raise ValueError('cluster_split_ratio debe pertenecer a (0,1]')
        if not 0 <= self.diversity_fraction <= 1 or self.diversity_cost_tolerance < 0:
            raise ValueError('Parámetros de diversidad inválidos')


# ======================== MÓDULO CONSERVADO: models.py ========================
# """Entidades del problema; sin dependencias de las rutinas de búsqueda."""
from dataclasses import dataclass, field
from typing import Optional


@dataclass(frozen=True)
class Depot:
    id: str
    capacity: float
    opening_cost: float  # equivalente diario COP, NUNCA inversión bruta
    x: Optional[float] = None
    y: Optional[float] = None
    input_cost: float = 0.0
    input_cost_basis: str = 'investment'


@dataclass(frozen=True)
class Client:
    id: str
    delivery: float
    pickup: float
    x: Optional[float] = None
    y: Optional[float] = None


@dataclass
class Instance:
    depots: dict[str, Depot]
    clients: dict[str, Client]
    distance: dict[str, dict[str, float]]
    time: dict[str, dict[str, float]]
    vehicle_capacity: float = 2500.0
    cost_per_km: float = 5000.0
    vehicle_rental_daily: float = 147572.55

    def d(self, a: str, b: str) -> float:
        return self.distance[a][b]

    def t(self, a: str, b: str) -> float:
        return self.time[a][b]


@dataclass
class Cluster:
    id: str
    depot_id: str
    kind: str   # DELIVERY o PICKUP
    clients: list[str]
    demand: float


@dataclass
class Segment:
    id: str
    cluster_id: str
    depot_id: str
    kind: str
    clients: list[str]
    demand: float
    start: str
    end: str
    internal_km: float


@dataclass
class Pairing:
    depot_id: str
    delivery_segment: Optional[str]
    pickup_segment: Optional[str]
    separate_km: float
    combined_km: float
    saving_km: float


@dataclass
class Stop:
    order: int
    node_id: str
    operation: str
    amount: float
    load_before: float
    load_after: float
    leg_km: float
    leg_minutes: float
    cumulative_km: float
    cumulative_minutes: float


@dataclass
class Route:
    id: str
    depot_id: str
    delivery_segment: Optional[str]
    pickup_segment: Optional[str]
    stops: list[Stop]
    km: float
    minutes: float
    cost: float
    initial_load: float


@dataclass
class Solution:
    id: str
    config_id: str
    start_number: int
    seed: int
    open_depots: list[str]
    assignment: dict[str, str]
    clusters: list[Cluster]
    segments: list[Segment]
    pairings: list[Pairing]
    routes: list[Route]
    opening_cost: float
    transport_cost: float
    total_cost: float
    distance_km: float
    feasible: bool = False
    violations: list[str] = field(default_factory=list)
    vehicle_rental_cost: float = 0.0  # total diario de vehículos (una vez por ruta)


@dataclass
class CombinationLog:
    config_id: str
    depots: list[str]
    capacity: float
    opening_cost: float
    accepted: bool
    reason: str
    status: str = 'PENDIENTE'
    attempts: int = 0
    best_total_cost: Optional[float] = None
    prune_incumbent_cost: Optional[float] = None


@dataclass
class AttemptLog:
    solution_id: str
    config_id: str
    start_number: int
    seed: int
    feasible: bool
    open_depots: list[str]
    opening_cost: float = 0.0
    transport_cost: float = 0.0
    total_cost: float = 0.0
    km: float = 0.0
    routes: int = 0
    reason: str = ""
    elapsed_seconds: float = 0.0
    phase: str = ''


@dataclass
class RunResult:
    lower_bound: int
    upper_bound: int
    total_combinations: int
    combinations: list[CombinationLog]
    attempts: list[AttemptLog]
    best_by_config: dict[str, Solution]
    best_global: Optional[Solution]
    alternatives: list[Solution]
    settings: dict[str, object]
    evolution: list[dict[str, object]] = field(default_factory=list)
    elapsed_seconds: float = 0.0
    stop_reason: str = ''


# ======================== MÓDULO CONSERVADO: io_data.py ========================
# """Lectura segura de Excel, identificadores y matrices dirigidas."""
import math
import pandas as pd


def normalize_id(value: object) -> str:
    if pd.isna(value):
        raise ValueError('Identificador vacío')
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        if math.isfinite(float(value)) and float(value).is_integer():
            return str(int(value))
    s = str(value).strip()
    if not s or s.lower() in ('nan', 'none', '<na>'):
        raise ValueError('Identificador vacío')
    return s


def _number(value: object, context: str, positive: bool = False) -> float:
    try:
        v = float(value)
    except (TypeError, ValueError) as e:
        raise ValueError(f'{context}: valor no numérico: {value!r}') from e
    if not math.isfinite(v) or v < 0 or (positive and v <= 0):
        raise ValueError(f'{context}: se esperaba valor finito {"positivo" if positive else "no negativo"}: {v}')
    return v


def _coordinates(row: pd.Series, cx: str, cy: str, where: str) -> tuple[float | None, float | None]:
    if cx not in row.index and cy not in row.index:
        return None, None
    if cx not in row.index or cy not in row.index:
        raise ValueError(f'{where}: debe haber ambas columnas de coordenadas')
    x, y = row[cx], row[cy]
    if pd.isna(x) and pd.isna(y):
        return None, None
    return _signed(x, f'{where} x'), _signed(y, f'{where} y')


def _signed(value: object, context: str) -> float:
    try:
        val = float(value)
    except (ValueError, TypeError) as e:
        raise ValueError(f'{context}: coordenada no numérica') from e
    if not math.isfinite(val):
        raise ValueError(f'{context}: coordenada no finita')
    return val


def _read_matrix(path: str, sheet: str, row_name: str, node_set: set[str], what: str):
    df = pd.read_excel(path, sheet_name=sheet, dtype=object, engine='openpyxl')
    if row_name not in df.columns:
        raise ValueError(f'{what}: columna de etiquetas {row_name!r} no encontrada')
    cols = [normalize_id(c) for c in df.columns if c != row_name]
    rows = [normalize_id(x) for x in df[row_name]]
    if len(set(cols)) != len(cols) or len(set(rows)) != len(rows):
        raise ValueError(f'{what}: identificadores duplicados en filas o columnas')
    if set(cols) != node_set or set(rows) != node_set:
        raise ValueError(f'{what}: nodos de filas faltantes={sorted(node_set - set(rows))}, sobrantes={sorted(set(rows) - node_set)}; columnas faltantes={sorted(node_set-set(cols))}, sobrantes={sorted(set(cols)-node_set)}')
    matrix = {}
    for _, row in df.iterrows():
        a = normalize_id(row[row_name])
        matrix[a] = {}
        for column, b in zip((c for c in df.columns if c != row_name), cols):
            v = _number(row[column], f'{what}[{a},{b}]')
            if a == b and abs(v) > 1e-8:
                raise ValueError(f'{what}: diagonal no cero en {a}: {v}')
            matrix[a][b] = v
    return matrix


def read_instance(cfg: InputConfig, s: SearchConfig) -> Instance:
    s.validate()
    if cfg.depot_cost_basis not in ('investment', 'daily'):
        raise ValueError('depot_cost_basis debe ser investment o daily')
    if cfg.depot_useful_life_years <= 0 or cfg.days_per_year <= 0:
        raise ValueError('Vida útil y días/año deben ser positivos')
    dep_df = pd.read_excel(cfg.depots_file, sheet_name=cfg.depots_sheet, dtype=object, engine='openpyxl')
    cli_df = pd.read_excel(cfg.clients_file, sheet_name=cfg.clients_sheet, dtype=object, engine='openpyxl')
    for col in (cfg.depot_id, cfg.depot_capacity, cfg.depot_cost):
        if col not in dep_df.columns:
            raise ValueError(f'Bodegas: columna requerida ausente {col}')
    for col in (cfg.client_id, cfg.delivery, cfg.pickup):
        if col not in cli_df.columns:
            raise ValueError(f'Clientes: columna requerida ausente {col}')
    depots, clients = {}, {}
    for i, r in dep_df.iterrows():
        name = normalize_id(r[cfg.depot_id])
        if name in depots:
            raise ValueError(f'Bodega duplicada: {name}')
        x, y = _coordinates(r, cfg.depot_x, cfg.depot_y, f'Bodega {name}')
        input_cost = _number(r[cfg.depot_cost], f'{name} costo de instalación')
        daily_cost = (input_cost / (cfg.depot_useful_life_years * cfg.days_per_year)
                      if cfg.depot_cost_basis == 'investment' else input_cost)
        depots[name] = Depot(name, _number(r[cfg.depot_capacity], f'{name} capacidad', positive=True),
                              daily_cost, x, y, input_cost, cfg.depot_cost_basis)
    for i, r in cli_df.iterrows():
        name = normalize_id(r[cfg.client_id])
        if name in clients:
            raise ValueError(f'Cliente duplicado: {name}')
        x, y = _coordinates(r, cfg.client_x, cfg.client_y, f'Cliente {name}')
        delivery = _number(r[cfg.delivery], f'{name} entrega')
        pickup = _number(r[cfg.pickup], f'{name} recogida')
        if delivery + pickup <= 0:
            raise ValueError(f'Cliente {name}: ninguna operación requerida')
        clients[name] = Client(name, delivery, pickup, x, y)
    if not depots or not clients:
        raise ValueError('Se requieren al menos una bodega y un cliente')
    collision = set(depots) & set(clients)
    if collision:
        raise ValueError(f'IDs coincidentes entre bodegas y clientes: {sorted(collision)}')
    nodes = set(depots) | set(clients)
    distance = _read_matrix(cfg.distance_file, cfg.distance_sheet, cfg.matrix_row_id, nodes, 'Distancias')
    time = _read_matrix(cfg.time_file, cfg.time_sheet, cfg.matrix_row_id, nodes, 'Tiempos')
    instance = Instance(depots, clients, distance, time, s.vehicle_capacity,
                        s.cost_per_km, s.vehicle_rental_daily)
    validate_instance(instance)
    return instance


def validate_instance(inst: Instance) -> None:
    if inst.vehicle_capacity <= 0 or inst.cost_per_km < 0 or inst.vehicle_rental_daily < 0:
        raise ValueError('Parámetros de vehículos inválidos')
    total = sum(c.delivery for c in inst.clients.values())
    if total > sum(d.capacity for d in inst.depots.values()) + 1e-7:
        raise ValueError(f'Capacidad de bodegas insuficiente para demanda de entregas: {total}')
    for c in inst.clients.values():
        if c.delivery > inst.vehicle_capacity + 1e-7 or c.pickup > inst.vehicle_capacity + 1e-7:
            raise ValueError(f'Cliente {c.id}: demanda individual supera capacidad del vehículo')
        if c.delivery > max(d.capacity for d in inst.depots.values()) + 1e-7:
            raise ValueError(f'Cliente {c.id}: entrega superior a cualquier bodega')
    nodes = set(inst.depots) | set(inst.clients)
    for name, matrix in [('Distancias', inst.distance), ('Tiempos', inst.time)]:
        if set(matrix) != nodes:
            raise ValueError(f'{name}: filas no coinciden con nodos')
        for a in nodes:
            if set(matrix[a]) != nodes:
                raise ValueError(f'{name}: columnas faltantes para {a}')
            for b in nodes:
                _number(matrix[a][b], f'{name}[{a},{b}]')
                if a == b and abs(matrix[a][b]) > 1e-8:
                    raise ValueError(f'{name}: diagonal diferente de cero en {a}')


# ======================== MÓDULO CONSERVADO: construction.py ========================
# """Fase I: combinaciones, regret asignación, CFRS, Cheapest Insertion y matching."""
from itertools import combinations
import math
import random
from scipy.optimize import linear_sum_assignment
import numpy as np

EPS = 1e-7


class ConstructionError(Exception):
    """Fracaso de la construcción; NO demuestra infactibilidad matemática."""


def facility_bounds(ins: Instance) -> tuple[int, int]:
    demand = sum(c.delivery for c in ins.clients.values())
    capacity = [d.capacity for d in ins.depots.values()]
    if sum(capacity) + EPS < demand:
        raise ValueError('La capacidad total disponible es menor que la demanda de entrega')
    # Caso límite solo recogidas: se necesita al menos una bodega, aunque 0/Cap = 0.
    lower = max(1, math.ceil((demand - EPS) / max(capacity)))
    # El límite por la menor capacidad NO es una cota superior válida para
    # minimizar costo: pueden convenir más instalaciones cercanas y baratas.
    upper = len(ins.depots)
    return lower, upper


def biased_index(n: int, p: float, rng: random.Random) -> int:
    """Geométrica truncada: peso proporcional a p*(1-p)^r, r=0..n-1."""
    if n <= 1 or p >= 1:
        return 0
    r = rng.random()
    weights = [p * (1 - p) ** i for i in range(n)]
    norm = sum(weights)
    total = 0.0
    for i, weight in enumerate(weights):
        total += weight / norm
        if r < total:
            return i
    return n - 1


def _unrank(n: int, k: int, rank: int) -> tuple[int, ...]:
    """Combinación lexicográfica desde su rango, sin crear el espacio completo."""
    positions = []
    first = 0
    for remaining in range(k, 0, -1):
        for index in range(first, n - remaining + 1):
            count = math.comb(n - index - 1, remaining - 1)
            if rank < count:
                positions.append(index)
                first = index + 1
                break
            rank -= count
    return tuple(positions)


def generate_combinations(ins: Instance, s: SearchConfig, rng: random.Random):
    """Enumera cardinalidades desde la cota necesaria hasta TODAS las bodegas.

    La capacidad agregada es un filtro necesario, no una prueba de factibilidad.
    El costo de apertura no se usa todavía para podar: requiere un incumbente
    completo, factible y VALIDADO obtenido en la búsqueda.
    """
    lo, hi = facility_bounds(ins)
    ids = sorted(ins.depots)
    total = sum(math.comb(len(ids), k) for k in range(lo, hi + 1))
    records: list[CombinationLog] = []
    accepted: list[CombinationLog] = []
    demand = sum(c.delivery for c in ins.clients.values())
    for k in range(lo, hi + 1):
        for comb in combinations(ids, k):
            capacity = sum(ins.depots[d].capacity for d in comb)
            opening = sum(ins.depots[d].opening_cost for d in comb)
            valid_capacity = capacity + EPS >= demand
            rec = CombinationLog(
                f'CFG-{len(records)+1:05d}', list(comb), capacity, opening,
                valid_capacity,
                'Capacidad agregada suficiente (no garantiza asignación)' if valid_capacity
                else 'Descartada: capacidad agregada insuficiente',
                status='PENDIENTE' if valid_capacity else 'DESCARTADA_CAPACIDAD')
            records.append(rec)
            if valid_capacity:
                accepted.append(rec)
    assert len(records) == total
    return lo, hi, total, records, accepted


def proxy_cost(ins: Instance, depot: str, client: str) -> float:
    """Aproximación por viaje de ida y vuelta por cada tipo de operación."""
    c = ins.clients[client]
    trips = int(c.delivery > EPS) + int(c.pickup > EPS)
    return trips * (ins.d(depot, client) + ins.d(client, depot)) * ins.cost_per_km


def _can_fit_remaining(ins: Instance, remaining: list[str], residual: dict[str, float]) -> bool:
    if sum(residual.values()) + EPS < sum(ins.clients[c].delivery for c in remaining):
        return False
    return all(any(cap + EPS >= ins.clients[c].delivery for cap in residual.values()) for c in remaining)


def _repair_assign(ins: Instance, client: str, assignment: dict[str, str], residual: dict[str, float],
                   options: list[str], depth: int, visited: set[str]) -> bool:
    """Reubicaciones constructivas limitadas, con rollback completo."""
    need = ins.clients[client].delivery
    for depot in sorted(options, key=lambda d: proxy_cost(ins, d, client)):
        if residual[depot] + EPS >= need:
            assignment[client] = depot
            residual[depot] -= need
            return True
        if depth <= 0:
            continue
        occupied = [c for c, d in assignment.items() if d == depot and c not in visited and ins.clients[c].delivery > EPS]
        occupied.sort(key=lambda c: (ins.clients[c].delivery, -proxy_cost(ins, depot, c)), reverse=True)
        for victim in occupied:
            before_assign = assignment.copy()
            before_resid = residual.copy()
            load = ins.clients[victim].delivery
            del assignment[victim]
            residual[depot] += load
            alternatives = [d for d in options if d != depot]
            if _repair_assign(ins, victim, assignment, residual, alternatives, depth - 1, visited | {client, victim}):
                if residual[depot] + EPS >= need:
                    assignment[client] = depot
                    residual[depot] -= need
                    return True
            assignment.clear(); assignment.update(before_assign)
            residual.clear(); residual.update(before_resid)
    return False


def _assignment_options(ins: Instance, client: str, pending: list[str],
                        residual: dict[str, float]) -> list[tuple[float, str]]:
    """Costo de atención + oportunidad perdida para las operaciones pendientes.

    Ambas partes están en COP y se calculan con la misma aproximación vial.
    Se evita asignar si viola condiciones NECESARIAS de capacidad; no se
    declara infactibilidad por fallar una prueba de empaquetamiento voraz.
    """
    needed = ins.clients[client].delivery
    others = [c for c in pending if c != client]
    base = {}
    for other in others:
        demand = ins.clients[other].delivery
        choices = [proxy_cost(ins, d, other) for d, cap in residual.items() if cap + EPS >= demand]
        if not choices:
            return []
        base[other] = min(choices)
    options = []
    for depot, free in residual.items():
        if free + EPS < needed:
            continue
        after = residual.copy()
        after[depot] -= needed
        if not _can_fit_remaining(ins, others, after):
            continue
        opportunity = 0.0
        for other in others:
            demand = ins.clients[other].delivery
            # Únicamente cambia la elegibilidad de la bodega seleccionada.
            # Todas las alternativas se comparan con el mismo costo proxy.
            if after[depot] + EPS < demand and free + EPS >= demand:
                choices = [proxy_cost(ins, d, other) for d, cap in after.items()
                           if cap + EPS >= demand]
                if not choices:
                    break
                opportunity += max(0.0, min(choices) - base[other])
        else:
            options.append((proxy_cost(ins, depot, client) + opportunity, depot))
    return sorted(options)


def assign_clients(ins: Instance, open_depots: list[str], s: SearchConfig,
                   rng: random.Random) -> dict[str, str]:
    residual = {d: ins.depots[d].capacity for d in open_depots}
    clients = list(ins.clients)
    if len(open_depots) == 1:
        depot = open_depots[0]
        if sum(c.delivery for c in ins.clients.values()) > residual[depot] + EPS:
            raise ConstructionError('Bodega única: capacidad insuficiente')
        return {c: depot for c in clients}
    assigned: dict[str, str] = {}
    pending = [c for c in clients if ins.clients[c].delivery > EPS]
    while pending:
        ranked = []
        for client in pending:
            options = _assignment_options(ins, client, pending, residual)
            if not options:
                continue
            # Opción forzada por capacidad = prioridad explícita, no un costo
            # monetario ficticio. Con alternativas, ventaja RELATIVA calculada
            # entre costos estimados homogéneos (COP) y factibles.
            forced = len(options) == 1
            regret = ((options[1][0] - options[0][0]) / max(options[1][0], 1.0)
                      if len(options) > 1 else 0.0)
            ranked.append((-int(forced), -regret, -ins.clients[client].delivery,
                           client, options))
        if not ranked:
            victim = max(pending, key=lambda c: ins.clients[c].delivery)
            if not _repair_assign(ins, victim, assigned, residual, open_depots,
                                  s.assignment_repairs, {victim}):
                raise ConstructionError(f'No se pudo asignar {victim} pese a reparaciones limitadas')
            pending.remove(victim)
            continue
        ranked.sort(key=lambda x: x[:4])
        selected = ranked[biased_index(len(ranked), s.geometric_p, rng)]
        client, options = selected[3], selected[4]
        depot = options[biased_index(len(options), s.geometric_p, rng)][1]
        assigned[client] = depot
        residual[depot] -= ins.clients[client].delivery
        pending.remove(client)
    for client in sorted((c for c in clients if ins.clients[c].delivery <= EPS)):
        options = sorted(open_depots, key=lambda d: (proxy_cost(ins, d, client), d))
        assigned[client] = options[biased_index(len(options), s.geometric_p, rng)]
    return assigned


def _insertion_delta(ins: Instance, depot: str, kind: str,
                     seq: list[str], item: str, pos: int,
                     closed: bool = False) -> float:
    """Incremento dirigido; los extremos abiertos se mantienen libres.

    DELIVERY ancla únicamente la salida depósito->primer cliente.
    PICKUP ancla únicamente el último cliente->depósito.
    closed=True ancla ambos extremos para estimar el coste de una ruta sola.
    """
    previous = seq[pos - 1] if pos else (depot if closed or kind == 'DELIVERY' else None)
    following = seq[pos] if pos < len(seq) else (depot if closed or kind == 'PICKUP' else None)
    return ((ins.d(previous, item) if previous is not None else 0.0)
            + (ins.d(item, following) if following is not None else 0.0)
            - (ins.d(previous, following) if previous is not None and following is not None else 0.0))


def _solo_cluster_km(ins: Instance, depot: str, seq: list[str]) -> float:
    return (ins.d(depot, seq[0])
            + sum(ins.d(a,b) for a,b in zip(seq, seq[1:]))
            + ins.d(seq[-1], depot))


def cluster_operations(ins: Instance, depot: str, assigned: dict[str, str], kind: str,
                       p: float, rng: random.Random, first_index: int,
                       split_ratio: float = 0.80) -> list[Cluster]:
    amount = (lambda c: ins.clients[c].delivery) if kind == 'DELIVERY' else (lambda c: ins.clients[c].pickup)
    pending = {c for c, d in assigned.items() if d == depot and amount(c) > EPS}
    result = []
    while pending:
        # Mantener semillas lejanas/voluminosas y selección geométrica.
        seeds = sorted(pending, key=lambda c: (-(ins.d(depot,c)+ins.d(c,depot))
                                                * (1 + amount(c)/ins.vehicle_capacity), c))
        seed = seeds[biased_index(len(seeds), p, rng)]
        items = [seed]
        load = amount(seed)
        pending.remove(seed)
        while True:
            viable = [c for c in pending if load + amount(c) <= ins.vehicle_capacity + EPS]
            if not viable:
                break
            ranked = []
            for client in viable:
                # Estimación de recorrido de la futura ruta completa, contando
                # salida y retorno para no fragmentar grupos por proximidad sola.
                delta, pos = min((_insertion_delta(ins, depot, kind, items, client, i, closed=True), i)
                                 for i in range(len(items)+1))
                ranked.append((delta, -amount(client), client, pos))
            ranked.sort()
            # Filtrar CADA candidato antes del sorteo geométrico. Comparar
            # incrementos y nueva ruta independiente en km DIRIGIDOS; ambos
            # incluyen salida y regreso al depósito en la estimación cerrada.
            # El emparejamiento posterior utiliza segmentos abiertos reales.
            admissible = [entry for entry in ranked
                          if entry[0] < split_ratio *
                          (ins.d(depot, entry[2]) + ins.d(entry[2], depot)) - EPS]
            if not admissible:
                break
            chosen = admissible[biased_index(len(admissible), p, rng)]
            client, pos = chosen[2], chosen[3]
            items.insert(pos, client)
            load += amount(client)
            pending.remove(client)
        result.append(Cluster(f'CL-{first_index + len(result):05d}', depot, kind, items, load))
    return result


def cheapest_insertion(ins: Instance, cluster: Cluster) -> Segment:
    if not cluster.clients:
        raise ConstructionError('Clúster vacío')
    # Preservar la semilla aleatorizada del CFRS, pero evaluar en CADA paso
    # todos los clientes aún pendientes y TODAS las posiciones de inserción.
    seq = [cluster.clients[0]]
    pending = cluster.clients[1:].copy()
    original_rank = {client: pos for pos, client in enumerate(cluster.clients)}
    while pending:
        delta, _, pos, item = min(
            (_insertion_delta(ins, cluster.depot_id, cluster.kind, seq, item, pos),
             original_rank[item], pos, item)
            for item in pending for pos in range(len(seq)+1)
        )
        seq.insert(pos, item)
        pending.remove(item)
    length = sum(ins.d(a,b) for a,b in zip(seq, seq[1:]))
    return Segment(f'SG-{cluster.id[3:]}', cluster.id, cluster.depot_id, cluster.kind,
                   seq, cluster.demand, seq[0], seq[-1], length)


def solo_distance(ins: Instance, seg: Segment) -> float:
    d = seg.depot_id
    return ins.d(d, seg.start) + seg.internal_km + ins.d(seg.end, d)


def joint_distance(ins: Instance, linehaul: Segment, backhaul: Segment) -> float:
    d = linehaul.depot_id
    return (ins.d(d,linehaul.start) + linehaul.internal_km +
            ins.d(linehaul.end,backhaul.start) + backhaul.internal_km + ins.d(backhaul.end,d))


def match_segments(ins: Instance, depot: str, segments: list[Segment]) -> list[Pairing]:
    lh = [x for x in segments if x.depot_id == depot and x.kind == 'DELIVERY']
    bh = [x for x in segments if x.depot_id == depot and x.kind == 'PICKUP']
    m, n = len(lh), len(bh)
    if not m or not n:
        return [Pairing(depot, x.id if x.kind == 'DELIVERY' else None,
                        x.id if x.kind == 'PICKUP' else None, solo_distance(ins, x),
                        solo_distance(ins, x), 0.0) for x in lh+bh]
    # Matriz de ahorros incrementales; m+n filas/columnas con dummies de costo 0.
    cost = np.zeros((m+n,m+n), dtype=float)
    for i,l in enumerate(lh):
        for j,b in enumerate(bh):
            # Comparar costo TOTAL diario: ruta mixta = 1 arriendo, rutas separadas = 2.
            cost[i,j] = ((joint_distance(ins,l,b) - solo_distance(ins,l) - solo_distance(ins,b))
                         * ins.cost_per_km - ins.vehicle_rental_daily)
    rows, cols = linear_sum_assignment(cost)
    result = []
    used_l, used_b = set(), set()
    for i,j in zip(rows,cols):
        if i < m and j < n and cost[i,j] < -EPS:
            l,b = lh[i],bh[j]
            alone = solo_distance(ins,l)+solo_distance(ins,b)
            together = joint_distance(ins,l,b)
            result.append(Pairing(depot,l.id,b.id,alone,together,alone-together))
            used_l.add(i); used_b.add(j)
    for i,l in enumerate(lh):
        if i not in used_l:
            dist = solo_distance(ins,l)
            result.append(Pairing(depot,l.id,None,dist,dist,0))
    for j,b in enumerate(bh):
        if j not in used_b:
            dist = solo_distance(ins,b)
            result.append(Pairing(depot,None,b.id,dist,dist,0))
    return result


def build_routes(ins: Instance, pairings: list[Pairing], segments: list[Segment]) -> list[Route]:
    lookup = {s.id:s for s in segments}
    routes = []
    for index, pair in enumerate(pairings,1):
        l = lookup[pair.delivery_segment] if pair.delivery_segment else None
        b = lookup[pair.pickup_segment] if pair.pickup_segment else None
        operations = ([(c, 'DELIVERY', ins.clients[c].delivery) for c in l.clients] if l else []) + \
                     ([(c, 'PICKUP', ins.clients[c].pickup) for c in b.clients] if b else [])
        load = l.demand if l else 0.0
        initial_load = load
        if load > ins.vehicle_capacity + EPS or (b and b.demand > ins.vehicle_capacity + EPS):
            raise ConstructionError('Segmentos exceden capacidad de vehículo')
        km = minutes = 0.0
        stops = [Stop(0,pair.depot_id,'DEPOT_START',0,load,load,0,0,0,0)]
        previous = pair.depot_id
        for no,(node,operation,amount) in enumerate(operations,1):
            leg = ins.d(previous,node)
            travel = ins.t(previous,node)
            newload = load - amount if operation == 'DELIVERY' else load + amount
            if newload < -EPS or newload > ins.vehicle_capacity + EPS:
                raise ConstructionError('Carga no factible en reconstrucción')
            km += leg; minutes += travel
            stops.append(Stop(no,node,operation,amount,load,newload,leg,travel,km,minutes))
            previous,load = node,newload
        leg = ins.d(previous,pair.depot_id)
        travel = ins.t(previous,pair.depot_id)
        km += leg; minutes += travel
        stops.append(Stop(len(stops),pair.depot_id,'DEPOT_END',0,load,load,leg,travel,km,minutes))
        routes.append(Route(f'R-{index:05d}',pair.depot_id,pair.delivery_segment,
                            pair.pickup_segment,stops,km,minutes,km*ins.cost_per_km,initial_load))
    return routes


def construct_solution(ins: Instance, selected: CombinationLog, s: SearchConfig,
                       seed: int, start_number: int, solution_id: str) -> Solution:
    rng = random.Random(seed)
    assignment = assign_clients(ins, selected.depots, s, rng)
    clusters = []
    for depot in selected.depots:
        for kind in ('DELIVERY', 'PICKUP'):
            chunks = cluster_operations(ins,depot,assignment,kind,s.geometric_p,rng,len(clusters)+1,
                                        s.cluster_split_ratio)
            clusters.extend(chunks)
    segments = [cheapest_insertion(ins,c) for c in clusters]
    pairings = [p for depot in selected.depots for p in match_segments(ins,depot,segments)]
    routes = build_routes(ins,pairings,segments)
    opening = sum(ins.depots[x].opening_cost for x in selected.depots)
    transport = sum(r.cost for r in routes)
    total_km = sum(r.km for r in routes)
    rent = len(routes) * ins.vehicle_rental_daily
    return Solution(solution_id, selected.config_id, start_number, seed,
                    selected.depots.copy(), assignment, clusters, segments, pairings,
                    routes, opening, transport, opening+rent+transport, total_km,
                    vehicle_rental_cost=rent)


# ======================== MÓDULO CONSERVADO: validation.py ========================
# """Validador independiente: recompone cargas, arcos, coberturas y costos desde la instancia."""
from collections import Counter, defaultdict
import math


def _close(a, b, tol=1e-5):
    return math.isfinite(float(a)) and math.isfinite(float(b)) and abs(a-b) <= tol * max(1, abs(a), abs(b))


def validate_solution(ins: Instance, sol: Solution) -> list[str]:
    errors: list[str] = []

    def check(condition: bool, message: str):
        if not condition:
            errors.append(message)

    opened = sol.open_depots
    check(len(opened) == len(set(opened)) and all(d in ins.depots for d in opened), 'Bodegas abiertas desconocidas/duplicadas')
    check(set(sol.assignment) == set(ins.clients), 'Asignación: faltan/sobran clientes')
    for c,d in sol.assignment.items():
        check(d in opened, f'Cliente {c}: bodega asignada no abierta')
    for d in opened:
        assigned = sum(ins.clients[c].delivery for c,x in sol.assignment.items() if c in ins.clients and x == d)
        check(assigned <= ins.depots[d].capacity + EPS, f'Bodega {d}: capacidad de entregas excedida')

    # Particiones de clústeres y segmentos: validación sin confiar en el constructor.
    expected = Counter((c,'DELIVERY') for c,v in ins.clients.items() if v.delivery > EPS)
    expected.update((c,'PICKUP') for c,v in ins.clients.items() if v.pickup > EPS)
    cluster_counts = Counter()
    cluster_ids = set()
    cl_by_id = {}
    for cl in sol.clusters:
        check(cl.id not in cluster_ids, f'Clúster duplicado: {cl.id}')
        cluster_ids.add(cl.id)
        cl_by_id[cl.id] = cl
        check(cl.kind in ('DELIVERY','PICKUP'), f'{cl.id}: tipo no válido')
        check(cl.depot_id in opened, f'{cl.id}: bodega no abierta')
        demand = 0.0
        for c in cl.clients:
            if c not in ins.clients:
                errors.append(f'{cl.id}: cliente desconocido {c}')
                continue
            cluster_counts[c,cl.kind] += 1
            check(sol.assignment.get(c) == cl.depot_id, f'{cl.id}: cliente {c} no pertenece a bodega')
            amount = ins.clients[c].delivery if cl.kind == 'DELIVERY' else ins.clients[c].pickup
            check(amount > EPS, f'{cl.id}: operación inexistente de {c}')
            demand += amount
        check(_close(cl.demand,demand), f'{cl.id}: demanda de clúster incorrecta')
        check(demand <= ins.vehicle_capacity + EPS, f'{cl.id}: capacidad vehicular')
    check(cluster_counts == expected, f'Clústeres: coberturas erróneas {list((cluster_counts-expected).items())}, faltantes={list((expected-cluster_counts).items())}')

    seg_by_id = {}
    seg_counts = Counter()
    for seg in sol.segments:
        check(seg.id not in seg_by_id, f'Segmento duplicado {seg.id}')
        seg_by_id[seg.id] = seg
        cl = cl_by_id.get(seg.cluster_id)
        check(cl is not None, f'{seg.id}: clúster inexistente')
        if cl:
            check(seg.depot_id == cl.depot_id and seg.kind == cl.kind, f'{seg.id}: metadatos distintos al clúster')
            check(Counter(seg.clients) == Counter(cl.clients), f'{seg.id}: operaciones distintas al clúster')
        check(bool(seg.clients), f'{seg.id}: segmento vacío')
        if seg.clients:
            check(seg.start == seg.clients[0] and seg.end == seg.clients[-1], f'{seg.id}: extremos incorrectos')
        if all(c in ins.clients for c in seg.clients):
            for c in seg.clients:
                seg_counts[c,seg.kind] += 1
            demand = sum(ins.clients[c].delivery if seg.kind=='DELIVERY' else ins.clients[c].pickup for c in seg.clients)
            check(_close(demand,seg.demand), f'{seg.id}: demanda incorrecta')
            km = sum(ins.d(a,b) for a,b in zip(seg.clients,seg.clients[1:]))
            check(_close(km,seg.internal_km), f'{seg.id}: distancia interna incorrecta')
    check(seg_counts == expected, 'Segmentos: cobertura de operaciones diferente a la requerida')
    check(len(sol.segments) == len(sol.clusters), 'Cantidad de clústeres y segmentos diferente')

    paired = Counter()
    expected_route_pairs = Counter()
    for pair in sol.pairings:
        check(pair.delivery_segment is not None or pair.pickup_segment is not None, 'Emparejamiento vacío')
        included = []
        for sid,kind in ((pair.delivery_segment,'DELIVERY'),(pair.pickup_segment,'PICKUP')):
            if sid is None:
                continue
            paired[sid] += 1
            seg = seg_by_id.get(sid)
            check(seg is not None, f'Emparejamiento de segmento desconocido {sid}')
            if seg:
                check(seg.depot_id == pair.depot_id and seg.kind == kind, f'Segmento {sid}: bodega/tipo incorrecto al emparejar')
                included.append(seg)
        expected_route_pairs[pair.depot_id,pair.delivery_segment,pair.pickup_segment] += 1
        if included:
            if len(included) == 1:
                km = ins.d(pair.depot_id, included[0].start) + sum(ins.d(a,b) for a,b in zip(included[0].clients,included[0].clients[1:])) + ins.d(included[0].end,pair.depot_id)
                check(_close(pair.combined_km,km) and _close(pair.separate_km,km) and _close(pair.saving_km,0), 'Costo incorrecto del emparejamiento individual')
            else:
                l,b = included
                solo_l = ins.d(pair.depot_id,l.start) + sum(ins.d(a,z) for a,z in zip(l.clients,l.clients[1:])) + ins.d(l.end,pair.depot_id)
                solo_b = ins.d(pair.depot_id,b.start) + sum(ins.d(a,z) for a,z in zip(b.clients,b.clients[1:])) + ins.d(b.end,pair.depot_id)
                together = ins.d(pair.depot_id,l.start) + sum(ins.d(a,z) for a,z in zip(l.clients,l.clients[1:])) + ins.d(l.end,b.start) + sum(ins.d(a,z) for a,z in zip(b.clients,b.clients[1:])) + ins.d(b.end,pair.depot_id)
                check(_close(pair.separate_km,solo_l+solo_b) and _close(pair.combined_km,together) and _close(pair.saving_km,solo_l+solo_b-together), 'Ahorro de emparejamiento incorrecto')
                # Una ruta mixta puede recorrer más kilómetros y aun ahorrar
                # dinero diario por evitar un arriendo de vehículo.
                check((together-solo_l-solo_b)*ins.cost_per_km <= ins.vehicle_rental_daily+EPS,
                      'Emparejamiento más costoso que rutas independientes incluyendo arriendo')
    check(paired == Counter({s.id:1 for s in sol.segments}), 'No se utilizan exactamente una vez todos los segmentos')

    visit_counts = Counter()
    routes_pairs = Counter()
    total_km = 0.0
    total_transport = 0.0
    route_ids = set()
    for route in sol.routes:
        check(route.id not in route_ids, f'Ruta duplicada: {route.id}')
        route_ids.add(route.id)
        routes_pairs[route.depot_id,route.delivery_segment,route.pickup_segment] += 1
        check(route.depot_id in opened, f'Ruta {route.id}: bodega no abierta')
        if len(route.stops) < 3:
            errors.append(f'Ruta {route.id}: menos de 3 visitas incluyendo depósitos')
            continue
        first,last = route.stops[0],route.stops[-1]
        check(first.operation == 'DEPOT_START' and last.operation == 'DEPOT_END', f'{route.id}: falta salida o retorno')
        check(first.node_id == route.depot_id and last.node_id == route.depot_id, f'{route.id}: sale/regresa a bodega diferente')
        expected_ops = []
        for seg_id,kind in ((route.delivery_segment,'DELIVERY'),(route.pickup_segment,'PICKUP')):
            if seg_id:
                seg = seg_by_id.get(seg_id)
                if seg:
                    expected_ops.extend((c,kind) for c in seg.clients)
        actual_ops = [(st.node_id,st.operation) for st in route.stops[1:-1]]
        check(expected_ops == actual_ops, f'{route.id}: orden/operaciones distintos a segmentos')
        initial = sum(ins.clients[c].delivery for c,k in actual_ops if c in ins.clients and k=='DELIVERY')
        check(_close(initial,route.initial_load), f'{route.id}: carga inicial errónea')
        load = initial
        km = minutes = 0.0
        pickup_started = False
        for no, stop in enumerate(route.stops):
            check(stop.order == no, f'{route.id}: índice de visita erróneo en {no}')
            if no == 0:
                check(_close(stop.load_before,load) and _close(stop.load_after,load) and
                      _close(stop.leg_km,0) and _close(stop.leg_minutes,0) and
                      _close(stop.cumulative_km,0) and _close(stop.cumulative_minutes,0), f'{route.id}: origen incorrecto')
                continue
            previous = route.stops[no-1].node_id
            if previous not in ins.distance or stop.node_id not in ins.distance.get(previous,{}):
                errors.append(f'{route.id}: arco inexistente {previous}->{stop.node_id}')
                continue
            leg = ins.d(previous,stop.node_id)
            travel = ins.t(previous,stop.node_id)
            km += leg; minutes += travel
            check(_close(stop.leg_km,leg) and _close(stop.leg_minutes,travel), f'{route.id}: distancia/tiempo de arco incorrectos en visita {no}')
            check(_close(stop.cumulative_km,km) and _close(stop.cumulative_minutes,minutes), f'{route.id}: acumulados erróneos en {no}')
            check(_close(stop.load_before,load), f'{route.id}: carga anterior errónea en {no}')
            if no < len(route.stops)-1:
                c = ins.clients.get(stop.node_id)
                if not c:
                    errors.append(f'{route.id}: cliente desconocido {stop.node_id}')
                    continue
                check(sol.assignment.get(c.id) == route.depot_id, f'{route.id}: cliente asignado a otra bodega: {c.id}')
                if stop.operation == 'DELIVERY':
                    check(not pickup_started, f'{route.id}: entrega posterior a recogida')
                    amount = c.delivery
                    load -= amount
                elif stop.operation == 'PICKUP':
                    pickup_started = True
                    amount = c.pickup
                    load += amount
                else:
                    errors.append(f'{route.id}: operación interior inválida {stop.operation}')
                    continue
                visit_counts[c.id,stop.operation] += 1
                check(amount > EPS and _close(stop.amount,amount), f'{route.id}: demanda de operación incorrecta')
            else:
                check(stop.amount == 0, f'{route.id}: última visita tiene demanda')
            check(_close(stop.load_after,load), f'{route.id}: carga posterior errónea en {no}')
            check(-EPS <= load <= ins.vehicle_capacity+EPS, f'{route.id}: violación capacidad del vehículo')
        check(_close(route.km,km) and _close(route.minutes,minutes), f'{route.id}: distancia/tiempo total incorrectos')
        check(_close(route.cost,km*ins.cost_per_km), f'{route.id}: costo de transporte incorrecto')
        total_km += km
        total_transport += km*ins.cost_per_km
    check(routes_pairs == expected_route_pairs, 'Rutas y emparejamientos no coinciden (incluye duplicados)')
    for depot in opened:
        departing = sum(bool(r.stops) and r.stops[0].node_id == depot and
                        r.stops[0].operation == 'DEPOT_START' for r in sol.routes if r.depot_id == depot)
        returning = sum(bool(r.stops) and r.stops[-1].node_id == depot and
                        r.stops[-1].operation == 'DEPOT_END' for r in sol.routes if r.depot_id == depot)
        used = sum(r.depot_id == depot for r in sol.routes)
        check(departing == returning == used,
              f'Bodega {depot}: vehículos salen={departing}, regresan={returning}, usados={used}')
    check(visit_counts == expected, f'Visitas: omitidas={list((expected-visit_counts).items())}, duplicadas={list((visit_counts-expected).items())}')
    opening = sum(ins.depots[d].opening_cost for d in opened if d in ins.depots)
    check(_close(sol.opening_cost,opening), 'Costo equivalente diario de bodegas incorrecto')
    check(_close(sol.transport_cost,total_transport), 'Costo de transporte total incorrecto')
    check(_close(sol.distance_km,total_km), 'Kilometraje total incorrecto')
    rental = len(sol.routes) * ins.vehicle_rental_daily
    check(_close(sol.vehicle_rental_cost,rental), 'Arriendo diario de vehículos incorrecto')
    check(_close(sol.total_cost,opening+rental+total_transport),
          'Costo global diario incorrecto: bodegas + arriendo + transporte')
    return errors


# ======================== MÓDULO CONSERVADO: search.py ========================
# """Multiarranque, memoria de élite y terminación por construcciones completadas.
#
# Solo se poda mediante el costo obligatorio de apertura frente al costo COMPLETO
# validado del incumbente, siempre que transporte >= 0. No se utiliza distancia
# estimada como cota inferior ni se confunde fallo de construcción con infactibilidad.
# """
import copy
import math
import random
from dataclasses import fields, replace
from time import perf_counter


def signature(solution: Solution) -> tuple:
    """Estructura canónica, independiente del número de vehículo u orden de rutas."""
    routes = tuple(sorted((r.depot_id, tuple((s.node_id,s.operation) for s in r.stops)) for r in solution.routes))
    return tuple(sorted(solution.open_depots)), tuple(sorted(solution.assignment.items())), routes


def diversity(a: Solution, b: Solution) -> float:
    if set(a.open_depots) != set(b.open_depots):
        return 1.0
    clients = a.assignment.keys() | b.assignment.keys()
    diff_a = sum(a.assignment.get(c)!=b.assignment.get(c) for c in clients)/max(1,len(clients))
    route_a = {tuple((s.node_id,s.operation) for s in r.stops) for r in a.routes}
    route_b = {tuple((s.node_id,s.operation) for s in r.stops) for r in b.routes}
    diff_r = 1 - len(route_a&route_b)/max(1,len(route_a|route_b))
    return max(diff_a,diff_r)


class EliteMemory:
    def __init__(self, top: int, diversity_fraction: float, tolerance: float):
        self.top = top
        self.diversity_fraction = diversity_fraction
        self.tolerance = tolerance
        self.best_global: Solution | None = None
        self.best_by_config: dict[str,Solution] = {}
        self._pool: dict[tuple,Solution] = {}

    def add(self, candidate: Solution):
        if not candidate.feasible or candidate.violations:
            raise ValueError('No almacenar soluciones sin validar')
        if self.best_global is None or candidate.total_cost < self.best_global.total_cost - 1e-6:
            self.best_global = copy.deepcopy(candidate)
        previous = self.best_by_config.get(candidate.config_id)
        if previous is None or candidate.total_cost < previous.total_cost - 1e-6:
            self.best_by_config[candidate.config_id] = copy.deepcopy(candidate)
        key = signature(candidate)
        previous = self._pool.get(key)
        if previous is None or candidate.total_cost < previous.total_cost - 1e-6:
            self._pool[key] = copy.deepcopy(candidate)
        if len(self._pool) > max(self.top * 12, 50):
            best = sorted(self._pool.values(),key=lambda x:(x.total_cost,x.id))
            selected = best[:max(self.top*6,25)]
            self._pool = {signature(x):x for x in selected}

    def get_alternatives(self) -> list[Solution]:
        if not self._pool:
            return []
        candidates = sorted(self._pool.values(), key=lambda x:(x.total_cost,x.id))
        selected = [candidates[0]]
        deferred = []
        limit = candidates[0].total_cost * (1 + self.tolerance)
        for sol in candidates[1:]:
            if len(selected) >= self.top:
                break
            if sol.total_cost <= limit and all(diversity(sol,x) >= self.diversity_fraction for x in selected):
                selected.append(sol)
            else:
                deferred.append(sol)
        if len(selected) < self.top:
            for sol in sorted(deferred,key=lambda x:(x.total_cost,x.id)):
                selected.append(sol)
                if len(selected) >= self.top:
                    break
        return [copy.deepcopy(x) for x in sorted(selected,key=lambda x:(x.total_cost,x.id))]


def multi_start(ins: Instance, s: SearchConfig,
                planned_schedule: list[str] | None = None,
                on_feasible=None, deterministic_seen=None,
                explore_extended=False) -> RunResult:
    """Fase definitiva reproducible por número de construcciones; DOE fijo.

    Distribución adaptativa sin porcentajes: ensanchar progresivamente el número
    de configuraciones probadas según ceil(sqrt(arranques+1)) y, entre las ya
    exploradas, reforzar las no dominadas en (costo real, oportunidades recibidas).
    La selección de nuevas combinaciones usa únicamente el costo de apertura
    como orden de visita, no como estimación del transporte ni como infactibilidad.
    La duración se mide, pero NO interviene en ninguna decisión ni terminación.
    """
    s.validate()
    # Condición para poda matemáticamente válida, sin otros costos negativos.
    if ins.cost_per_km < 0 or any(d.opening_cost < 0 for d in ins.depots.values()) or any(
            value < 0 for row in ins.distance.values() for value in row.values()):
        raise ValueError('Poda por apertura requiere costos de apertura, distancia y transporte no negativos')
    started = perf_counter()
    rng = random.Random(s.seed)
    lo,hi,total,combos,accepted = generate_combinations(ins,s,rng)
    memory = EliteMemory(s.keep_top,s.diversity_fraction,s.diversity_cost_tolerance)
    attempts: list[AttemptLog] = []
    seen_structures: set[tuple] = set()  # medición de diversidad; NO cambia las decisiones
    evolution: list[dict] = []
    counter = 0

    # Mantener la trayectoria inicial de multiarranque original: primero
    # configuraciones dentro del intervalo antiguo, con EXACTAMENTE su
    # barajado y desempates. La cota ampliada no consume su presupuesto.
    legacy_upper = min(len(ins.depots), max(1, math.ceil(
        (sum(c.delivery for c in ins.clients.values()) - EPS) /
        min(d.capacity for d in ins.depots.values()))))
    legacy = [rec for rec in accepted if len(rec.depots) <= legacy_upper]
    extended = [rec for rec in accepted if len(rec.depots) > legacy_upper]
    rng.shuffle(legacy)
    legacy.sort(key=lambda rec: rec.opening_cost)
    # La proximidad sirve para ORDENAR (no para descartar) las instalaciones
    # adicionales. Operaciones vectorizadas evitan recalcular arcos repetidos.
    if extended:
        depot_access = {d: np.fromiter(
            (ins.d(d, c) + ins.d(c, d) for c in ins.clients), dtype=float)
            for d in ins.depots}
        def access(rec):
            return float(np.minimum.reduce(
                [depot_access[d] for d in rec.depots]).sum())
        extended.sort(key=lambda rec: (rec.opening_cost, access(rec)))
    if explore_extended and extended:
        # Tras el lote inicial, ensayar de forma controlada DOS combinaciones
        # con mayor cardinalidad entre las primeras configuraciones a explorar.
        # No añade construcciones ni modifica la política adaptativa existente.
        ordered = legacy[:3] + extended[:2] + legacy[3:] + extended[2:]
    else:
        ordered = legacy + extended

    def incumbent() -> float:
        return memory.best_global.total_cost if memory.best_global else math.inf

    def prunable(combo) -> bool:
        return combo.opening_cost >= incumbent()  # transporte no negativo

    def mark_pruned(combo):
        if combo.attempts == 0:
            combo.status = 'DESCARTADA_COSTO'
            combo.prune_incumbent_cost = incumbent()
            combo.reason = ('Apertura >= costo total de solución factible validada; '
                            'transporte >= 0; no puede mejorar al incumbente')

    def one_attempt(combo, phase: str) -> None:
        nonlocal counter
        next_number = combo.attempts + 1
        sid = f'SOL-{counter+1:06d}'
        seed = rng.randrange(2**63)
        log = AttemptLog(sid,combo.config_id,next_number,seed,False,
                         combo.depots.copy(),combo.opening_cost,phase=phase)
        # El reloj es exclusivamente observacional: no filtra ni inicia arranques.
        t0 = perf_counter()
        counter += 1
        combo.attempts = next_number
        try:
            # Híbrido únicamente para la búsqueda definitiva: PRIMER intento
            # de cada configuración determinista, TODOS los siguientes sesgados.
            # DOE 'fixed' conserva su p experimental sin alteraciones.
            # No repetir la primera construcción determinista de una misma
            # configuración y cierre en los lotes posteriores. Las siguientes
            # pruebas siguen siendo construcciones nuevas con su propia semilla.
            deterministic_key = (round(s.cluster_split_ratio, 12), tuple(combo.depots))
            deterministic_first = (s.execution_mode == 'budget' and next_number == 1
                and (deterministic_seen is None or deterministic_key not in deterministic_seen))
            attempt_cfg = replace(s, geometric_p=1.0) if deterministic_first else s
            if deterministic_first and deterministic_seen is not None:
                deterministic_seen.add(deterministic_key)
            candidate = construct_solution(ins,combo,attempt_cfg,seed,combo.attempts,sid)
            failures = validate_solution(ins,candidate)
            candidate.violations = failures
            candidate.feasible = not failures
            log.feasible = candidate.feasible
            log.transport_cost = candidate.transport_cost
            log.total_cost = candidate.total_cost
            log.km = candidate.distance_km
            log.routes = len(candidate.routes)
            log.reason = 'Validación correcta' if not failures else ' | '.join(failures)
            if candidate.feasible:
                struct = signature(candidate)
                log.is_novel = struct not in seen_structures
                seen_structures.add(struct)
                # Almacenar solamente costos reales recalculados y validados.
                before = incumbent()
                memory.add(candidate)
                if on_feasible is not None:
                    on_feasible(candidate)
                combo.best_total_cost = min(combo.best_total_cost or math.inf, candidate.total_cost)
                combo.status = 'EVALUADA_FACTIBLE'
                combo.reason = 'Al menos una construcción completa y validada'
                if incumbent() < before - 1e-6:
                    evolution.append(dict(solucion_id=sid,config_id=combo.config_id,
                                          arranque=combo.attempts,semilla=str(seed),
                                          segundos=perf_counter()-started,
                                          costo_apertura=candidate.opening_cost,
                                          costo_transporte=candidate.transport_cost,
                                          costo_total=candidate.total_cost,
                                          km=candidate.distance_km,
                                          rutas=len(candidate.routes)))
            elif combo.status != 'EVALUADA_FACTIBLE':
                combo.status = 'EVALUADA_SIN_SOLUCION'
                combo.reason = 'Construcción incumple validación; no demuestra infactibilidad'
        except ConstructionError as exc:
            log.reason = f'Construcción no completada (no prueba infactibilidad): {exc}'
            if combo.status != 'EVALUADA_FACTIBLE':
                combo.status = 'EVALUADA_SIN_SOLUCION'
                combo.reason = 'Falló construcción; la configuración no es matemáticamente infactible'
        except Exception as exc:
            raise RuntimeError(f'Error técnico en configuración {combo.config_id}, '
                               f'arranque {combo.attempts}, semilla {seed}: '
                               f'{type(exc).__name__}: {exc}') from exc
        finally:
            log.elapsed_seconds = perf_counter() - t0
            attempts.append(log)

    if s.execution_mode == 'fixed':
        # DOE: cantidad de intentos idéntica para cada tratamiento bajo una
        # semilla (sin poda por incumbente que sesgaría el presupuesto).
        chosen = rng.sample(accepted, min(len(accepted), s.combinations))
        for combo in chosen:
            for _ in range(s.starts_per_combination):
                one_attempt(combo,'PRESUPUESTO_FIJO')
        stop_reason = 'Presupuesto fijo de construcciones completado'
    else:
        # Producción: MISMO número de construcciones, independiente de la CPU.
        # No existe cuota porcentual para exploración/intensificación.
        index = 0

        def next_unexplored():
            nonlocal index
            while index < len(ordered):
                combo = ordered[index]
                index += 1
                if prunable(combo):
                    mark_pruned(combo)
                    continue
                return combo
            return None

        def best_revisit():
            """Frente no dominado (costo REAL validado, nº de oportunidades).

            Todos los desempates usan identificadores estables; nunca duración.
            """
            observed = [c for c in ordered if c.attempts > 0 and not prunable(c)]
            feasible = [c for c in observed if c.config_id in memory.best_by_config]
            if feasible:
                frontier = [c for c in feasible if not any(
                    (other.attempts <= c.attempts and
                     other.best_total_cost <= c.best_total_cost and
                     (other.attempts < c.attempts or
                      other.best_total_cost < c.best_total_cost))
                    for other in feasible if other is not c)]
                return min(frontier, key=lambda c:(c.attempts,c.best_total_cost,c.config_id))
            return min(observed,key=lambda c:(c.attempts,c.opening_cost,c.config_id)) if observed else None

        if planned_schedule is not None:
            if len(planned_schedule) != s.construction_budget:
                raise ValueError('El control pareado requiere una agenda de exactamente '
                                 'construction_budget configuraciones')
            available = {c.config_id: c for c in accepted}
            missing = sorted(set(planned_schedule) - set(available))
            if missing:
                raise ValueError(f'Agenda de control contiene configuraciones inválidas: {missing}')

        while counter < s.construction_budget:
            # Control experimental pareado: reutilizar EXACTAMENTE el orden de
            # configuraciones y semillas internas de la búsqueda híbrida; no
            # cambiar la distribución ni reconstruir configuraciones nuevas.
            # Los costos reales NO cambian la agenda previamente fijada.
            if planned_schedule is not None:
                one_attempt(available[planned_schedule[counter]], 'CONTROL_AGENDA_IDÉNTICA')
                continue
            # Referencias puramente deterministas: no gastar presupuesto en
            # reconstrucciones idénticas de la misma configuración. Se ensaya
            # otra configuración válida, manteniendo el mismo presupuesto.
            if s.geometric_p == 1.0:
                # En la referencia, cada construcción se aplica a una bodega
                # distinta; sin réplicas de una construcción determinista.
                # No se poda por apertura durante este CONTROL experimental:
                # así se conserva exactamente el mismo nº de evaluaciones.
                next_combo = ordered[index] if index < len(ordered) else None
                if next_combo is None:
                    break
                index += 1
                one_attempt(next_combo, 'REFERENCIA_DETERMINISTA')
                continue
            explored_count = sum(c.attempts > 0 for c in ordered)
            next_count = counter + 1
            desired_explored = math.isqrt(next_count)
            if desired_explored**2 < next_count:
                desired_explored += 1
            seek_new = (memory.best_global is None or explored_count < desired_explored)
            combo = next_unexplored() if seek_new else None
            phase = 'EXPLORACION_DETERMINISTA' if combo is not None else 'INTENSIFICACION_SESGADA'
            if combo is None:
                combo = best_revisit()
                if combo is None:
                    combo = next_unexplored()
                    phase = 'EXPLORACION_DETERMINISTA'
            if combo is None:
                break
            if combo.attempts == 0:
                phase = 'EXPLORACION_DETERMINISTA'
            else:
                phase = 'INTENSIFICACION_SESGADA'
            one_attempt(combo, phase)

        stop_reason = ('Presupuesto fijo de ' + str(s.construction_budget) +
                       ' construcciones completado' if counter == s.construction_budget
                       else 'Sin configuraciones pendientes ni revisitables que puedan mejorar')
        for combo in accepted:
            if combo.attempts == 0:
                if prunable(combo):
                    mark_pruned(combo)
                else:
                    combo.status = 'PENDIENTE_PRESUPUESTO'
                    combo.reason = 'Sin construcción dentro del presupuesto fijo; no infactible'
            elif prunable(combo) and combo.status == 'EVALUADA_SIN_SOLUCION':
                combo.reason += '; ahora su apertura no puede mejorar al incumbente'

    elapsed = perf_counter() - started
    params = {f.name:getattr(s,f.name) for f in fields(s)}
    if s.execution_mode == 'budget':
        inspected = sum(c.attempts > 0 for c in accepted)
        warning = (f'Presupuesto de {s.construction_budget} construcciones: '
                   f'{len(attempts)} completas y {inspected} configuraciones '
                   'exploradas. Ajuste construction_budget manualmente si requiere '
                   'más diversidad.') if len(attempts) < 2 or inspected < min(2,len(accepted)) else ''
    else:
        warning = ''
    params.update(construcciones_estructuralmente_distintas=len(seen_structures),
                  construcciones_duplicadas=sum(a.feasible for a in attempts)-len(seen_structures),
                  advertencia_presupuesto=warning,segundos_reales=elapsed,causa_finalizacion=stop_reason,
                  construcciones_configuradas=s.construction_budget if s.execution_mode=='budget' else
                    s.combinations*s.starts_per_combination,
                  construcciones_completadas=len(attempts),
                  criterio_parada='construcciones' if s.execution_mode=='budget' else 'DOE_fijo',
                  politica_hibrida=('control determinista con agenda fija' if planned_schedule is not None else
                      f'primera determinista por CFG (p=1); siguientes sesgadas (p={s.geometric_p})') if s.execution_mode=='budget' else 'p del DOE',
                  configuraciones_construidas=sum(c.attempts>0 for c in accepted),
                  configuraciones_pendientes=sum(c.status=='PENDIENTE_PRESUPUESTO' for c in accepted),
                  configuraciones_podadas_costo=sum(c.status=='DESCARTADA_COSTO' for c in accepted))
    return RunResult(lo,hi,total,combos,attempts,copy.deepcopy(memory.best_by_config),
                     copy.deepcopy(memory.best_global),memory.get_alternatives(),params,
                     evolution,elapsed,stop_reason)


# ======================== MÓDULO CONSERVADO: excel_io.py ========================
# """Archivo único Excel bidireccional para conservar soluciones completas de Fase I.
# No utiliza JSON ni pickle como formato de intercambio.
# """
from collections import defaultdict
from pathlib import Path
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Esquemas fijos, aun cuando una hoja esté vacía.
COLUMNS = {
    'Resumen': ['indicador','valor'],
    'Cotas': ['tipo','valor','descripcion'],
    'Combinaciones': ['config_id','bodegas','cantidad','capacidad','costo_apertura','aceptada','estado','arranques','mejor_costo','costo_global_para_poda','motivo'],
    'Arranques': ['estructura_nueva','solucion_id','config_id','arranque','semilla','fase','segundos','factible','bodegas','costo_apertura','costo_transporte','costo_total','km','rutas','motivo'],
    'Evolucion': ['solucion_id','config_id','arranque','semilla','segundos','costo_apertura','costo_transporte','costo_total','km','rutas'],
    'MejoresConfig': ['config_id','solucion_id','costo_total','km','rutas','bodegas'],
    'Alternativas': ['posicion','solucion_id','costo_total','diferencia_mejor','bodegas','rutas','km'],
    'Soluciones': ['solucion_id','config_id','arranque','semilla','es_global','en_alternativas','mejor_config','factible','costo_apertura','costo_arriendo_vehiculos','costo_transporte','costo_total','km','rutas','vehiculos_utilizados','validacion'],
    'BodegasAbiertas': ['solucion_id','bodega_id','capacidad','costo_apertura','costo_diario_bodega','costo_entrada','unidad_costo_entrada','entrega_asignada','recogida_asignada','vehiculos_usados','vehiculos_salen','vehiculos_regresan','arriendo_vehiculos_diario','km_diarios','costo_transporte_diario','costo_total_diario'],
    'Asignaciones': ['solucion_id','cliente_id','bodega_id','entrega','recogida'],
    'Clusters': ['solucion_id','cluster_id','bodega_id','tipo','demanda','operaciones'],
    'MiembrosCluster': ['solucion_id','cluster_id','orden','cliente_id'],
    'Segmentos': ['solucion_id','segmento_id','cluster_id','bodega_id','tipo','demanda','inicio','fin','km_internos'],
    'VisitasSegmento': ['solucion_id','segmento_id','orden','cliente_id'],
    'Emparejamientos': ['solucion_id','bodega_id','segmento_entrega','segmento_recogida','km_separados','km_combinados','ahorro_km'],
    'Rutas': ['solucion_id','ruta_id','bodega_id','segmento_entrega','segmento_recogida','carga_inicial','km','minutos','costo','costo_transporte_diario','arriendo_vehiculo_diario','costo_total_ruta_diario','vehiculo_sale','vehiculo_regresa'],
    'Visitas': ['solucion_id','ruta_id','orden','nodo_id','operacion','cantidad','carga_antes','carga_despues','km_arco','min_arco','km_acum','min_acum'],
    'Validacion': ['solucion_id','valida','incumplimientos'],
    'Parametros': ['parametro','valor'],
}


def _solutions(result: RunResult) -> list[Solution]:
    allsol = {}
    for sol in list(result.best_by_config.values()) + result.alternatives + ([result.best_global] if result.best_global else []):
        allsol[sol.id] = sol
    return sorted(allsol.values(), key=lambda sol:(sol.total_cost,sol.id))


def export_results_excel(path: str, ins: Instance, result: RunResult,
                         experiment_sheets: dict[str, list[dict]] | None = None) -> str:
    rows: dict[str,list[dict]] = {sheet:[] for sheet in COLUMNS}
    def add(sheet, **kw):
        rows[sheet].append(kw)

    feasible_count = sum(a.feasible for a in result.attempts)
    best=result.best_global
    info = {
        'numero_bodegas_candidatas':len(ins.depots), 'numero_clientes':len(ins.clients),
        'demanda_entregas':sum(c.delivery for c in ins.clients.values()),
        'demanda_recogidas':sum(c.pickup for c in ins.clients.values()),
        'capacidad_vehiculo':ins.vehicle_capacity,'costo_cop_km':ins.cost_per_km,
        'base_costo_bodegas_entrada':next(iter(ins.depots.values())).input_cost_basis,
        'vida_util_instalaciones_anios':15,
        'dias_por_anio':365,
        'arriendo_diario_por_vehiculo_cop':ins.vehicle_rental_daily,
        'costo_diario_bodegas_cop':best.opening_cost if best else None,
        'costo_diario_arriendo_cop':best.vehicle_rental_cost if best else None,
        'costo_diario_transporte_cop':best.transport_cost if best else None,
        'costo_total_diario_cop':best.total_cost if best else None,
        'vehiculos_diarios_utilizados':len(best.routes) if best else None,
        'cota_inferior_bodegas':result.lower_bound,'cota_superior_bodegas':result.upper_bound,
        'total_combinaciones_intervalo':result.total_combinations,
        'combinaciones_generadas':len(result.combinations),
        'combinaciones_evaluadas':sum(x.attempts > 0 for x in result.combinations),
        'combinaciones_pendientes':sum(x.status == 'PENDIENTE_PRESUPUESTO' for x in result.combinations),
        'combinaciones_podadas_costo':sum(x.status == 'DESCARTADA_COSTO' for x in result.combinations),
        'tiempo_real_segundos':result.elapsed_seconds,
        'causa_finalizacion':result.stop_reason,
        'combinaciones_aceptadas':sum(c.accepted for c in result.combinations),
        'combinaciones_descartadas':sum(not c.accepted for c in result.combinations),
        'arranques_ejecutados':len(result.attempts),'arranques_factibles':feasible_count,
        'construcciones_totales_busqueda_multi': result.settings.get('construcciones_totales_multi',len(result.attempts)),
        'construcciones_factibles_busqueda_multi': result.settings.get('construcciones_factibles_multi',feasible_count),
        'ejecuciones_independientes':result.settings.get('numero_ejecuciones_independientes',1),
        'mejor_solucion_id':best.id if best else 'Ninguna',
        'mejor_costo_total_cop':best.total_cost if best else None,
        'mejor_km':best.distance_km if best else None,
        'mejor_numero_rutas':len(best.routes) if best else None,
        'soluciones_completas_exportadas':len(_solutions(result)),
        'nota': 'Cota superior de instalaciones es exploratoria, no cota del costo óptimo. Capacidad de bodega contabiliza solo entregas.',
    }
    for key,value in info.items(): add('Resumen',indicador=key,valor=value)
    add('Cotas',tipo='Inferior',valor=result.lower_bound,descripcion='ceil(demanda entrega / capacidad máxima); mínimo una bodega')
    add('Cotas',tipo='Superior',valor=result.upper_bound,descripcion='min(N,ceil(demanda entrega / capacidad mínima)); exploración heurística')
    add('Cotas',tipo='Espacio',valor=result.total_combinations,descripcion='Combinaciones de tamaños en el intervalo')
    for c in result.combinations:
        add('Combinaciones',config_id=c.config_id,bodegas=' | '.join(c.depots),cantidad=len(c.depots),capacidad=c.capacity,
            costo_apertura=c.opening_cost,aceptada=c.accepted,estado=c.status,
            arranques=c.attempts,mejor_costo=c.best_total_cost,
            costo_global_para_poda=c.prune_incumbent_cost,motivo=c.reason)
    for a in result.attempts:
        add('Arranques',solucion_id=a.solution_id,config_id=a.config_id,arranque=a.start_number,
            semilla=str(a.seed),fase=a.phase,segundos=a.elapsed_seconds,
            factible=a.feasible,bodegas=' | '.join(a.open_depots),
            costo_apertura=a.opening_cost,costo_transporte=a.transport_cost,costo_total=a.total_cost,
            km=a.km,rutas=a.routes,motivo=a.reason,estructura_nueva=getattr(a,'is_novel',None))
    for event in result.evolution:
        add('Evolucion',**event)
    for cfg,s in sorted(result.best_by_config.items()):
        add('MejoresConfig',config_id=cfg,solucion_id=s.id,costo_total=s.total_cost,km=s.distance_km,
            rutas=len(s.routes),bodegas=' | '.join(s.open_depots))
    for idx,s in enumerate(result.alternatives,1):
        add('Alternativas',posicion=idx,solucion_id=s.id,costo_total=s.total_cost,
            diferencia_mejor=s.total_cost-best.total_cost if best else 0,
            bodegas=' | '.join(s.open_depots),rutas=len(s.routes),km=s.distance_km)
    alt_ids={x.id for x in result.alternatives}
    config_ids={x.id for x in result.best_by_config.values()}
    for s in _solutions(result):
        failures=validate_solution(ins,s)
        if failures:
            raise ValueError(f'No se exporta {s.id}: validación fallida: {failures[:6]}')
        add('Soluciones',solucion_id=s.id,config_id=s.config_id,arranque=s.start_number,
            semilla=str(s.seed),es_global=bool(best and s.id==best.id),en_alternativas=s.id in alt_ids,
            mejor_config=s.id in config_ids,factible=True,costo_apertura=s.opening_cost,
            costo_arriendo_vehiculos=s.vehicle_rental_cost,
            costo_transporte=s.transport_cost,costo_total=s.total_cost,km=s.distance_km,
            rutas=len(s.routes),vehiculos_utilizados=len(s.routes),validacion='OK')
        add('Validacion',solucion_id=s.id,valida=True,incumplimientos='')
        for depot in s.open_depots:
            fleet=[r for r in s.routes if r.depot_id==depot]
            departures=sum(r.stops[0].node_id==depot and r.stops[0].operation=='DEPOT_START' for r in fleet)
            arrivals=sum(r.stops[-1].node_id==depot and r.stops[-1].operation=='DEPOT_END' for r in fleet)
            if departures != arrivals or departures != len(fleet):
                raise AssertionError(f'Balance vehicular no válido en {depot}')
            depot_km=sum(r.km for r in fleet)
            daily_depot=ins.depots[depot].opening_cost
            daily_rent=len(fleet)*ins.vehicle_rental_daily
            daily_transport=sum(r.cost for r in fleet)
            add('BodegasAbiertas',solucion_id=s.id,bodega_id=depot,
                capacidad=ins.depots[depot].capacity,costo_apertura=daily_depot,
                costo_diario_bodega=daily_depot,
                costo_entrada=ins.depots[depot].input_cost,
                unidad_costo_entrada=ins.depots[depot].input_cost_basis,
                entrega_asignada=sum(ins.clients[c].delivery for c,d in s.assignment.items() if d==depot),
                recogida_asignada=sum(ins.clients[c].pickup for c,d in s.assignment.items() if d==depot),
                vehiculos_usados=len(fleet),vehiculos_salen=departures,
                vehiculos_regresan=arrivals,arriendo_vehiculos_diario=daily_rent,
                km_diarios=depot_km,costo_transporte_diario=daily_transport,
                costo_total_diario=daily_depot+daily_rent+daily_transport)
        for c,d in sorted(s.assignment.items()):
            add('Asignaciones',solucion_id=s.id,cliente_id=c,bodega_id=d,
                entrega=ins.clients[c].delivery,recogida=ins.clients[c].pickup)
        for cl in s.clusters:
            add('Clusters',solucion_id=s.id,cluster_id=cl.id,bodega_id=cl.depot_id,tipo=cl.kind,
                demanda=cl.demand,operaciones=len(cl.clients))
            for idx,c in enumerate(cl.clients,1):
                add('MiembrosCluster',solucion_id=s.id,cluster_id=cl.id,orden=idx,cliente_id=c)
        for seg in s.segments:
            add('Segmentos',solucion_id=s.id,segmento_id=seg.id,cluster_id=seg.cluster_id,
                bodega_id=seg.depot_id,tipo=seg.kind,demanda=seg.demand,
                inicio=seg.start,fin=seg.end,km_internos=seg.internal_km)
            for idx,c in enumerate(seg.clients,1):
                add('VisitasSegmento',solucion_id=s.id,segmento_id=seg.id,orden=idx,cliente_id=c)
        for p in s.pairings:
            add('Emparejamientos',solucion_id=s.id,bodega_id=p.depot_id,
                segmento_entrega=p.delivery_segment,segmento_recogida=p.pickup_segment,
                km_separados=p.separate_km,km_combinados=p.combined_km,ahorro_km=p.saving_km)
        for r in s.routes:
            add('Rutas',solucion_id=s.id,ruta_id=r.id,bodega_id=r.depot_id,
                segmento_entrega=r.delivery_segment,segmento_recogida=r.pickup_segment,
                carga_inicial=r.initial_load,km=r.km,minutos=r.minutes,costo=r.cost,
                costo_transporte_diario=r.cost,arriendo_vehiculo_diario=ins.vehicle_rental_daily,
                costo_total_ruta_diario=r.cost+ins.vehicle_rental_daily,
                vehiculo_sale=bool(r.stops and r.stops[0].node_id==r.depot_id),
                vehiculo_regresa=bool(r.stops and r.stops[-1].node_id==r.depot_id))
            for st in r.stops:
                add('Visitas',solucion_id=s.id,ruta_id=r.id,orden=st.order,nodo_id=st.node_id,
                    operacion=st.operation,cantidad=st.amount,carga_antes=st.load_before,
                    carga_despues=st.load_after,km_arco=st.leg_km,min_arco=st.leg_minutes,
                    km_acum=st.cumulative_km,min_acum=st.cumulative_minutes)
    for k,v in result.settings.items():
        add('Parametros',parametro=k,valor=v)
    if not hasattr(path, 'write'):
        Path(path).parent.mkdir(parents=True,exist_ok=True)
    with pd.ExcelWriter(path,engine='openpyxl') as writer:
        for name,cols in COLUMNS.items():
            df=pd.DataFrame(rows[name],columns=cols)
            df.to_excel(writer,sheet_name=name,index=False)
        # La calibración se adjunta al MISMO libro sin alterar el formato de
        # intercambio ni la reconstrucción de las soluciones de la Fase II.
        for name, data in (experiment_sheets or {}).items():
            if name in COLUMNS or len(name) > 31:
                raise ValueError(f'Nombre inválido/duplicado de hoja DOE: {name}')
            pd.DataFrame(data).to_excel(writer,sheet_name=name,index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes='A2'
            ws.auto_filter.ref=ws.dimensions
            ws.sheet_view.showGridLines=False
            ws.row_dimensions[1].height=26
            for cell in ws[1]:
                cell.fill=PatternFill('solid',fgColor='17365D')
                cell.font=Font(color='FFFFFF',bold=True,size=10)
                cell.alignment=Alignment(vertical='center',wrap_text=True)
            for index,col in enumerate(ws.columns,1):
                length=max((len(str(c.value or '')) for c in list(col)[:250]),default=12)
                ws.column_dimensions[get_column_letter(index)].width=min(48,max(13,length+2))
    return path


def _text(value):
    if pd.isna(value):
        return None
    return str(value).strip()


def _boolean(value):
    return str(value).lower() in ('true','1','sí','si') or value is True


def read_results_excel(path: str, instance: Instance | None = None):
    """Reconstruye todas las soluciones exportadas; con instancia, las revalida.

    Devuelve diccionario: solutions, best_global, alternatives, best_by_config.
    Los valores de nodos se leen como texto, incluso si parecen numéricos.
    """
    sheets = pd.read_excel(path,sheet_name=None,engine='openpyxl',dtype=str)
    missing = set(COLUMNS)-set(sheets)
    if missing:
        raise ValueError(f'Excel de salida incompleto: hojas ausentes {sorted(missing)}')
    def rows(sheet):
        return sheets[sheet].where(pd.notna(sheets[sheet]),None).to_dict('records')
    group = {}
    for name in ('BodegasAbiertas','Asignaciones','Clusters','MiembrosCluster',
                 'Segmentos','VisitasSegmento','Emparejamientos','Rutas','Visitas'):
        data=defaultdict(list)
        for row in rows(name):
            data[row['solucion_id']].append(row)
        group[name]=data

    solutions = {}
    for entry in rows('Soluciones'):
        sid=entry['solucion_id']
        depots=[r['bodega_id'] for r in group['BodegasAbiertas'][sid]]
        assignment={r['cliente_id']:r['bodega_id'] for r in group['Asignaciones'][sid]}
        members=defaultdict(list)
        for r in group['MiembrosCluster'][sid]:
            members[r['cluster_id']].append((int(r['orden']),r['cliente_id']))
        clusters=[Cluster(r['cluster_id'],r['bodega_id'],r['tipo'],
                          [c for _,c in sorted(members[r['cluster_id']])],float(r['demanda']))
                  for r in group['Clusters'][sid]]
        segment_members=defaultdict(list)
        for r in group['VisitasSegmento'][sid]:
            segment_members[r['segmento_id']].append((int(r['orden']),r['cliente_id']))
        segments=[Segment(r['segmento_id'],r['cluster_id'],r['bodega_id'],r['tipo'],
                          [c for _,c in sorted(segment_members[r['segmento_id']])],
                          float(r['demanda']),r['inicio'],r['fin'],float(r['km_internos']))
                  for r in group['Segmentos'][sid]]
        pairings=[Pairing(r['bodega_id'],_text(r['segmento_entrega']),_text(r['segmento_recogida']),
                          float(r['km_separados']),float(r['km_combinados']),float(r['ahorro_km']))
                  for r in group['Emparejamientos'][sid]]
        visits=defaultdict(list)
        for r in group['Visitas'][sid]:
            visits[r['ruta_id']].append(Stop(int(r['orden']),r['nodo_id'],r['operacion'],
                float(r['cantidad']),float(r['carga_antes']),float(r['carga_despues']),
                float(r['km_arco']),float(r['min_arco']),float(r['km_acum']),float(r['min_acum'])))
        routes=[Route(r['ruta_id'],r['bodega_id'],_text(r['segmento_entrega']),
                     _text(r['segmento_recogida']),sorted(visits[r['ruta_id']],key=lambda v:v.order),
                     float(r['km']),float(r['minutos']),float(r['costo']),float(r['carga_inicial']))
                for r in group['Rutas'][sid]]
        sol=Solution(sid,entry['config_id'],int(entry['arranque']),int(entry['semilla']),
                     depots,assignment,clusters,segments,pairings,routes,
                     float(entry['costo_apertura']),float(entry['costo_transporte']),
                     float(entry['costo_total']),float(entry['km']),_boolean(entry['factible']),
                     vehicle_rental_cost=float(entry['costo_arriendo_vehiculos']))
        if instance is not None:
            failures=validate_solution(instance,sol)
            if failures:
                raise ValueError(f'Solución importada {sid} viola restricciones: {failures[:8]}')
        solutions[sid]=sol
    ids={r['solucion_id'] for r in rows('Soluciones')}
    best_ids=[r['solucion_id'] for r in rows('Soluciones') if _boolean(r['es_global'])]
    if len(best_ids)>1:
        raise ValueError('Excel inconsistente: más de una mejor solución global')
    best=solutions[best_ids[0]] if best_ids else None
    alternatives=[solutions[r['solucion_id']] for r in sorted(rows('Alternativas'),key=lambda r:int(r['posicion']))]
    by_config={r['config_id']:solutions[r['solucion_id']] for r in rows('MejoresConfig')}
    for sid in ids:
        if sid not in solutions:
            raise ValueError(f'ID faltante {sid}')
    if best and (not alternatives or all(x.id!=best.id for x in alternatives)):
        raise ValueError('Excel inconsistente: mejor global fuera de alternativas')
    return {'solutions':solutions, 'best_global':best, 'alternatives':alternatives,
            'best_by_config':by_config}


# ======================== MÓDULO CONSERVADO: plotting.py ========================
# """Mapa con coordenadas disponibles, o diagrama esquemático sin georreferenciar."""
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np


def plot_solution(ins: Instance, sol: Solution, target: str) -> str:
    Path(target).parent.mkdir(parents=True,exist_ok=True)
    used = {st.node_id for r in sol.routes for st in r.stops}
    nodes = {**ins.depots,**ins.clients}
    geographic = all(nodes[node].x is not None and nodes[node].y is not None for node in used)
    fig,ax = plt.subplots(figsize=(13, max(6, min(20, 1.0*len(sol.routes)+3)) if not geographic else 10))
    palette = plt.get_cmap('tab20')
    if geographic:
        for depot in sol.open_depots:
            dd = ins.depots[depot]
            ax.scatter(dd.x,dd.y,marker='s',s=170,c='#222222',zorder=5)
            ax.annotate(f'Bodega {depot}',(dd.x,dd.y),xytext=(5,7),textcoords='offset points',weight='bold')
        for c in ins.clients.values():
            if c.id in used:
                ax.scatter(c.x,c.y,s=18,c='#999999',zorder=2)
        for ridx,r in enumerate(sol.routes):
            color = palette(ridx%20)
            for i, (a,b) in enumerate(zip(r.stops,r.stops[1:])):
                na,nb = nodes[a.node_id],nodes[b.node_id]
                kind=b.operation
                ax.plot([na.x,nb.x],[na.y,nb.y],ls='--' if kind=='PICKUP' else '-',
                        color=color,lw=1.8,alpha=.8,label=f'{r.id} ({r.depot_id})' if i==0 else None)
                midx,midy = (na.x+nb.x)/2,(na.y+nb.y)/2
                ax.annotate('→',(midx,midy),color=color,fontsize=10,ha='center',va='center')
            for stop in r.stops[1:-1]:
                node=nodes[stop.node_id]
                ax.annotate(f'{r.id[-2:]}:{stop.order}-{stop.node_id}',(node.x,node.y),
                            xytext=(4,4),textcoords='offset points',fontsize=7,color=color)
        ax.set_xlabel('Coordenada X (tal como aparece en Excel)')
        ax.set_ylabel('Coordenada Y (tal como aparece en Excel)')
        ax.set_title('Rutas: segmentos rectos entre coordenadas, NO geometría vial real')
    else:
        width=max((len(r.stops)-1 for r in sol.routes),default=2)
        for ridx,r in enumerate(sol.routes):
            y = len(sol.routes)-1-ridx
            color=palette(ridx%20)
            xs=np.linspace(0,width,max(len(r.stops),2))
            for i, stop in enumerate(r.stops):
                x=xs[i]
                if stop.operation.startswith('DEPOT'):
                    ax.scatter(x,y,marker='s',s=90,color=color,zorder=4)
                    label=f'{r.depot_id}\n{"Salida" if i==0 else "Retorno"}'
                else:
                    ax.scatter(x,y,marker='o' if stop.operation=='DELIVERY' else '^',
                               s=55,color=color,zorder=4)
                    label=f'{stop.order}. {stop.node_id}\n{"E" if stop.operation=="DELIVERY" else "R"} {stop.amount:g}'
                ax.annotate(label,(x,y),xytext=(0,11),textcoords='offset points',
                            ha='center',va='bottom',fontsize=8)
                if i>0:
                    previous=r.stops[i-1]
                    ax.annotate('',xy=(x-0.05,y),xytext=(xs[i-1]+0.05,y),
                                arrowprops={'arrowstyle':'->','color':color,'lw':2,
                                            'linestyle':'dashed' if stop.operation=='PICKUP' else 'solid'})
            ax.text(-0.7,y,f'{r.id}\n{r.km:.1f} km',ha='right',va='center',fontsize=8)
        ax.set_xlim(-1.5,width+0.8)
        ax.set_ylim(-0.8,len(sol.routes)+0.2)
        ax.set_xticks([]); ax.set_yticks([])
        ax.set_title('ESQUEMA de operaciones por ruta (NO representa posiciones ni trazados viales)')
        ax.spines[:].set_visible(False)
    ax.grid(alpha=.12)
    if geographic and len(sol.routes)<=15:
        ax.legend(loc='upper left',bbox_to_anchor=(1,1),fontsize=8)
    fig.suptitle(f'LogiData • {sol.id} • {len(sol.routes)} rutas • ${sol.total_cost:,.0f} COP',fontweight='bold',fontsize=13)
    fig.tight_layout()
    fig.savefig(target,dpi=160,bbox_inches='tight')
    plt.close(fig)
    return target


# ==================== NUEVA CAPA EXPERIMENTAL (SIN CAMBIAR CONSTRUCCIÓN) ====================
# Configuración de los Excel ORIGINALES: cambiar aquí únicamente si sus rutas cambian.
# Los argumentos CLI tienen prioridad. Las cuatro entradas deben existir en el
# mismo directorio, o bien indicar rutas completas mediante CLI.
from pathlib import Path as _Path
from collections import Counter as _Counter
from statistics import mean as _mean, stdev as _stdev
from dataclasses import replace as _replace
import argparse as _argparse
import os as _os
import hashlib as _hashlib
import csv as _csv

FILES_DEFAULT = {
    'bodegas': 'Localización de Instalaciones Proyecto.xlsx',
    'clientes': 'Instancia_01 (1).xlsx',
    'distancias': 'Distancias_LogiData(1).xlsx',
    'tiempos': 'Tiempos_LogiData(1).xlsx',
}
# Semillas elegidas de ANTEMANO, nunca ajustadas según el costo obtenido.
BUDGET_SEEDS = (5201, 5202, 5203, 5204, 5205, 5206, 5207, 5208, 5209, 5210)
CALIBRATION_SEEDS = (6101, 6102, 6103, 6104, 6105)
HOLDOUT_SEEDS = (7101, 7102, 7103, 7104, 7105)
BIAS_LEVELS = (0.40, 0.65, 0.80, 0.95)
SPLIT_LEVELS = (0.65, 0.80, 0.95)
# Aumenta desde 4 * 3 = 12 hasta 36 construcciones HÍBRIDAS por semilla y
# tratamiento. La confirmación de tratamientos finalistas se hace con 72/100.
DOE_CALIBRATION_BUDGET = 36


def _resolve_file(value: str) -> str:
    candidate = _Path(value).expanduser()
    choices = [candidate] if candidate.is_absolute() else [
        _Path.cwd()/candidate, _Path(__file__).resolve().parent/candidate,
        _Path('/mnt/data')/candidate,
    ]
    for x in choices:
        if x.is_file():
            return str(x.resolve())
    raise FileNotFoundError('No se encuentra ' + value +
          '. Coloca los 4 Excel junto al .py o especifica --bodegas etc.')


def setup_instance(filepaths=None, depot_cost_basis='investment',
                   vehicle_rental_daily=147572.55):
    f = {**FILES_DEFAULT, **(filepaths or {})}
    # Adaptación exclusiva de lectura: las instancias 05–25 incluyen las
    # cuatro fuentes en un único Excel. No cambian los objetos ni los algoritmos.
    rutas = {k: _resolve_file(f[k]) for k in
             ('bodegas', 'clientes', 'distancias', 'tiempos')}
    libro_unico = len(set(rutas.values())) == 1
    cfg = InputConfig(
        depots_file=rutas['bodegas'],
        clients_file=rutas['clientes'],
        distance_file=rutas['distancias'],
        time_file=rutas['tiempos'],
        depots_sheet='Hoja1', clients_sheet='Sheet1',
        distance_sheet='Distancias' if libro_unico else 'Distancias_km',
        time_sheet='Tiempos' if libro_unico else 'Tiempos_min',
        depot_id='ID', depot_capacity='Capacidad (Unidades)',
        depot_cost='Costo Fijo Apertura ($ COP)',
        depot_cost_basis=depot_cost_basis, client_id='Cliente',
        delivery='Demanda de Entrega', pickup='Demanda de Recogida',
        matrix_row_id='ID' if libro_unico else 'Origen / Destino',
        depot_x='Longitud', depot_y='Latitud',
        client_x='Longitud', client_y='Latitud')
    s = SearchConfig(execution_mode='budget', construction_budget=72,
        geometric_p=0.95, cluster_split_ratio=0.65,  # fallback: _cli recupera selección Excel validada
        seed=4110, keep_top=10, cost_per_km=5000.0, vehicle_capacity=2500.0,
        vehicle_rental_daily=vehicle_rental_daily)
    ins = read_instance(cfg, s)
    if len(ins.depots)==15 and len(ins.clients)==80:
        assert abs(sum(c.delivery for c in ins.clients.values())-15814)<1e-6
        assert abs(sum(c.pickup for c in ins.clients.values())-17329)<1e-6
    print(f'Datos: {len(ins.depots)} bodegas, {len(ins.clients)} clientes; '
          f'entregas={sum(c.delivery for c in ins.clients.values()):.0f}, '
          f'recogidas={sum(c.pickup for c in ins.clients.values()):.0f}',flush=True)
    return ins,s,cfg


def _structural_fingerprint(result):
    """Huella lógica: NO incluye tiempos, fechas ni datos binarios de Excel."""
    # Sin hash() de Python (sal aleatoria por proceso), usando SHA256 canónico.
    import json
    def sig(sol):
        if sol is None: return None
        return [round(sol.total_cost,6), sorted(sol.open_depots),
            sorted(sol.assignment.items()),
            sorted((r.depot_id, [(st.node_id,st.operation,st.amount,
                                   st.load_before,st.load_after) for st in r.stops])
                   for r in sol.routes)]
    obj={'attempts':[(a.config_id,a.start_number,int(a.seed),a.feasible,
                        round(a.total_cost,6) if a.total_cost is not None else None)
                      for a in result.attempts],
         'best':sig(result.best_global),
         'combos':[(x.config_id,x.attempts,x.status) for x in result.combinations]}
    return _hashlib.sha256(json.dumps(obj,sort_keys=True,ensure_ascii=True).encode()).hexdigest()


def _summarize_run(result,seed,budget,tag):
    best=result.best_global
    return dict(etapa=tag,semilla=int(seed),presupuesto=int(budget),
       construcciones=len(result.attempts),
       factibles=sum(a.feasible for a in result.attempts),
       configuraciones=sum(c.attempts>0 for c in result.combinations),
       pendientes=sum(c.status=='PENDIENTE_PRESUPUESTO' for c in result.combinations),
       costo_total=best.total_cost if best else None,
       costo_apertura=best.opening_cost if best else None,
       costo_transporte=best.transport_cost if best else None,
       km=best.distance_km if best else None,
       rutas=len(best.routes) if best else None,
       bodegas=' | '.join(sorted(best.open_depots)) if best else '',
       tiempo_seg=result.elapsed_seconds,
       distintas=result.settings.get('construcciones_estructuralmente_distintas'),
       repetidas=result.settings.get('construcciones_duplicadas'),
       huella=_structural_fingerprint(result))


def _aggregate_budget(rows,budget):
    subset=[r for r in rows if r['presupuesto']==budget]
    costs=[r['costo_total'] for r in subset if r['costo_total'] is not None]
    return dict(presupuesto=budget,n=len(subset),
        ejecuciones_con_solucion=len(costs),
        costo_min=min(costs) if costs else None,
        costo_medio=_mean(costs) if costs else None,
        costo_sd=_stdev(costs) if len(costs)>1 else 0.0,
        configuraciones_medias=_mean(r['configuraciones'] for r in subset),
        construcciones_medias=_mean(r['construcciones'] for r in subset),
        factibles_total=sum(r['factibles'] for r in subset),
        intentos_total=sum(r['construcciones'] for r in subset),
        rutas_medias=_mean(r['rutas'] for r in subset if r['rutas'] is not None),
        segundos_medios=_mean(r['tiempo_seg'] for r in subset))


def _paired_budget_test(rows, budget_a=72, budget_b=100):
    a={r['semilla']:r for r in rows if r['presupuesto']==budget_a}
    b={r['semilla']:r for r in rows if r['presupuesto']==budget_b}
    records=[]
    for seed in sorted(set(a)&set(b)):
        x,y=a[seed],b[seed]
        records.append(dict(semilla=seed,costo_72=x['costo_total'],costo_100=y['costo_total'],
            ahorro_100=(x['costo_total']-y['costo_total'] if x['costo_total'] is not None and y['costo_total'] is not None else None),
            segundos_extra=y['tiempo_seg']-x['tiempo_seg'],
            configuraciones_72=x['configuraciones'],configuraciones_100=y['configuraciones'],
            factibles_72=x['factibles'],factibles_100=y['factibles']))
    return records


def _budget_decision(pairs):
    """Criterio registrado previamente: evidencia de ahorro medio pareado.

    100 si el intervalo de confianza bilateral t 95% para (C72-C100) queda
    completamente por encima de cero y ambos presupuestos producen solución.
    De lo contrario 72 y se reporta que no se ha demostrado beneficio del costo.
    Evita introducir una tolerancia económica o un puntaje de pesos arbitrarios.
    """
    from scipy.stats import t as student_t
    diffs=[r['ahorro_100'] for r in pairs if r['ahorro_100'] is not None]
    if len(diffs)<2:
        return 72, 'Datos pareados insuficientes; conservar 72', None
    avg=_mean(diffs); sd=_stdev(diffs)
    half=float(student_t.ppf(.975,len(diffs)-1))*sd/math.sqrt(len(diffs))
    low,high=avg-half,avg+half
    choice=100 if low>0 else 72
    reason=(f'Ahorro medio pareado {avg:,.2f} COP; IC95% [{low:,.2f}, {high:,.2f}]. '
            + ('Ahorro medio positivo respaldado; seleccionar 100.' if low>0
               else 'No está demostrado ahorro medio positivo; conservar 72.'))
    return choice,reason,(low,high)


def compare_budgets(ins,base,seeds=BUDGET_SEEDS,verbose=True):
    """Diez semillas PREFIJADAS; presupuestos pareados y misma trayectoria inicial.

    En una semilla, los primeros 72 intentos de la ejecución de 100 son
    idénticos a los de 72 si datos, código, parámetros y entorno son iguales.
    """
    results=[]; tracking=[]; best_result=None
    saved_solutions={}  # (semilla, presupuesto) -> deepcopy de la mejor solución validada
    for seed in seeds:
        prefix=None
        for budget in (72,100):
            cfg=_replace(base, seed=seed, execution_mode='budget',
                         construction_budget=budget)
            result=multi_start(ins,cfg)
            r=_summarize_run(result,seed,budget,'Comparacion presupuesto')
            if prefix is None:
                prefix=[(a.config_id,a.start_number,a.seed,a.total_cost,a.feasible)
                        for a in result.attempts]
            else:
                current=[(a.config_id,a.start_number,a.seed,a.total_cost,a.feasible)
                        for a in result.attempts[:len(prefix)]]
                if current!=prefix:
                    raise AssertionError(f'La trayectoria de 72 no es prefijo de 100 (semilla {seed})')
            results.append(r)
            if result.best_global is not None:
                saved_solutions[(seed,budget)]=copy.deepcopy(result.best_global)
            if verbose:
                print(f'Semilla {seed}, {budget} construcciones: costo={r["costo_total"]}, '
                      f'validas={r["factibles"]}, configs={r["configuraciones"]}, '
                      f'seg={r["tiempo_seg"]:.2f}',flush=True)
            # Una solución para exportar; no se altera ninguna de las búsquedas.
            if result.best_global is not None and (best_result is None or
                 result.best_global.total_cost<best_result.best_global.total_cost):
                best_result=result
            if budget==100:
                tracking.append(dict(semilla=seed,nueva_construcciones=28,
                    costo_72=results[-2]['costo_total'],costo_100=r['costo_total'],
                    coste_reducido=(results[-2]['costo_total']-r['costo_total']
                        if results[-2]['costo_total'] is not None and r['costo_total'] is not None else None)))
    pairs=_paired_budget_test(results)
    choice,why,ci=_budget_decision(pairs)
    summaries=[_aggregate_budget(results,72),_aggregate_budget(results,100)]
    print('DECISIÓN PRESUPUESTO:',choice,why,flush=True)
    return dict(rows=results,paired=pairs,summary=summaries,choice=choice,
                reason=why,extra=tracking,saved_solutions=saved_solutions),best_result


def _exper_eval(ins,base,seed,p,ratio,budget,stage):
    cfg=_replace(base, seed=seed, geometric_p=p,cluster_split_ratio=ratio,
                 construction_budget=budget,execution_mode='budget')
    result=multi_start(ins,cfg)
    row=_summarize_run(result,seed,budget,stage)
    row.update(geometric_p=p,cluster_split_ratio=ratio,
               fase=('REFERENCIA APARTE: una configuración distinta por intento; '
                     'sin revisitaciones; NO es control pareado' if p == 1.0 else
                     'HÍBRIDA: primera determinista por configuración, siguientes sesgadas'))
    return row


def _doe_summary(rows,p,ratio):
    sub=[r for r in rows if r['geometric_p']==p and r['cluster_split_ratio']==ratio]
    good=[r for r in sub if r['costo_total'] is not None]
    values=[r['costo_total'] for r in good]
    return dict(geometric_p=p,cluster_split_ratio=ratio,semillas=len(sub),
                ejecuciones_con_solucion=len(good),
                construcciones_factibles=sum(r['factibles'] for r in sub),
                construcciones_totales=sum(r['construcciones'] for r in sub),
                costo_min=min(values) if values else None,
                costo_medio=_mean(values) if values else None,
                costo_sd=_stdev(values) if len(values)>1 else None,
                segundos_medios=_mean(r['tiempo_seg'] for r in sub))


def _doe_order(x):
    # Sin penalizaciones monetarias ficticias ni 1%: factibilidad, media,
    # desviación, tiempo. Preferir determinista solamente si su costo lo avala.
    return (-x['ejecuciones_con_solucion'],-x['construcciones_factibles'],
            x['costo_medio'] if x['costo_medio'] is not None else math.inf,
            x['costo_sd'] if x['costo_sd'] is not None else math.inf,
            x['segundos_medios'],x['geometric_p'],x['cluster_split_ratio'])



def _choose_hybrid_on_training(rows, ordered):
    """Comparar los dos mejores híbridos con semillas PAREADAS del entrenamiento.

    Si el IC95% de su diferencia media contiene cero, el costo no distingue
    claramente ambas alternativas: preferir menor desviación observada,
    luego menor tiempo. Nunca examina el holdout para hacer esta selección.
    """
    from scipy.stats import t as student_t
    finalists=[x for x in ordered if x['geometric_p']<1.0]
    first=finalists[0]
    if len(finalists)<2:
        return (first['geometric_p'],first['cluster_split_ratio']), 'Solo un tratamiento híbrido'
    second=finalists[1]
    if (first['ejecuciones_con_solucion']!=second['ejecuciones_con_solucion'] or
            first['construcciones_factibles']!=second['construcciones_factibles']):
        return (first['geometric_p'],first['cluster_split_ratio']), 'Factibilidad superior'
    fa={r['semilla']:r['costo_total'] for r in rows
        if r['geometric_p']==first['geometric_p'] and r['cluster_split_ratio']==first['cluster_split_ratio']}
    fb={r['semilla']:r['costo_total'] for r in rows
        if r['geometric_p']==second['geometric_p'] and r['cluster_split_ratio']==second['cluster_split_ratio']}
    diffs=[fb[k]-fa[k] for k in sorted(set(fa)&set(fb)) if fa[k] is not None and fb[k] is not None]
    if len(diffs)>=2:
        ave=_mean(diffs);std=_stdev(diffs)
        half=float(student_t.ppf(.975,len(diffs)-1))*std/math.sqrt(len(diffs))
        if ave-half <= 0 <= ave+half:
            chosen=min((first,second),key=lambda x:(x['costo_sd'] if x['costo_sd'] is not None else math.inf,
                                                 x['segundos_medios'],x['costo_medio']))
            return ((chosen['geometric_p'],chosen['cluster_split_ratio']),
                     f'IC95% del ahorro pareado [{ave-half:,.2f}, {ave+half:,.2f}] COP incluye cero; '
                     'menor dispersión entre tratamientos similares')
    return (first['geometric_p'],first['cluster_split_ratio']), 'Menor costo medio; diferencia no ambigua en comparación pareada'


def calibrate_hybrid(ins,base,verification_budget=72,
                     calibration_budget=DOE_CALIBRATION_BUDGET,
                     seeds=CALIBRATION_SEEDS,holdout=HOLDOUT_SEEDS):
    """DOE 4x3 híbrido + 3 referencias deterministas NO pareadas.

    Calibración: mismo número de construcciones por tratamiento (36 por semilla
    inicialmente). Posteriormente top 2 y determinista con presupuesto final
    seleccionado (72/100), usando semillas NUEVAS no empleadas para elegir.
    """
    from itertools import product
    combos=list(product(BIAS_LEVELS,SPLIT_LEVELS)) + [(1.0,r) for r in SPLIT_LEVELS]
    rows=[]
    for p,ratio in combos:
        for seed in seeds:
            r=_exper_eval(ins,base,seed,p,ratio,calibration_budget,'DOE calibracion')
            rows.append(r)
        agg=_doe_summary(rows,p,ratio)
        print(f'DOE p={p}, split={ratio}: costo_medio={agg["costo_medio"]}, '
              f'factibles={agg["construcciones_factibles"]}/{agg["construcciones_totales"]}',flush=True)
    summaries=[_doe_summary(rows,p,r) for p,r in combos]
    if not any(x['ejecuciones_con_solucion'] for x in summaries):
        raise RuntimeError('Ningún tratamiento completó una solución factible')
    ordered=sorted(summaries,key=_doe_order)
    # Pre-registrar finalistas: dos mejores + la determinista más competitiva.
    finalists=[(x['geometric_p'],x['cluster_split_ratio']) for x in
               [y for y in ordered if y['geometric_p']<1.0][:2]]
    deterministic=next((x for x in ordered if x['geometric_p']==1.0),None)
    if deterministic is not None:
        option=(deterministic['geometric_p'],deterministic['cluster_split_ratio'])
        if option not in finalists:
            finalists.append(option)
    holdout_rows=[]
    for p,ratio in finalists:
        for seed in holdout:
            r=_exper_eval(ins,base,seed,p,ratio,verification_budget,'DOE validacion independiente')
            holdout_rows.append(r)
        print(f'DOE HOLDOUT p={p}, split={ratio}: '+str(_doe_summary(holdout_rows,p,ratio)['costo_medio']),flush=True)
    holdout_summary=[_doe_summary(holdout_rows,p,r) for p,r in finalists]
    # Selección parametrizada solo mediante calibración; holdout confirma,
    # no se utiliza para ajustar ex post las semillas o elegir un ganador.
    # Elegir SOLO mediante el entrenamiento, antes de consultar el holdout.
    # Los 12 tratamientos híbridos compiten; p=1 es solo referencia.
    selected,selection_reason=_choose_hybrid_on_training(rows,ordered)
    print('SELECCIÓN DOE:',selected,selection_reason, '(holdout separado)',flush=True)
    effects=[]
    for name,levels,key in [('Geometrico',BIAS_LEVELS,'geometric_p'),
                            ('Cierre',SPLIT_LEVELS,'cluster_split_ratio')]:
        for lev in levels:
            subs=[r['costo_medio'] for r in summaries
                  if r[key]==lev and r['geometric_p']<1.0 and r['costo_medio'] is not None]
            effects.append(dict(factor=name,nivel=lev,
                                costo_medio_celdas=_mean(subs) if subs else None,
                                celdas=len(subs)))
    interaction=[]
    for ratio in SPLIT_LEVELS:
        base_cost=next((r['costo_medio'] for r in summaries
                        if r['geometric_p']==BIAS_LEVELS[0] and r['cluster_split_ratio']==ratio),None)
        for p in BIAS_LEVELS:
            row=next(r for r in summaries if r['geometric_p']==p and r['cluster_split_ratio']==ratio)
            interaction.append(dict(sesgo=p,cierre=ratio,costo_medio=row['costo_medio'],
                                    diferencia_vs_sesgo_040=(row['costo_medio']-base_cost if
                                         row['costo_medio'] is not None and base_cost is not None else None)))
    return dict(rows=rows,summary=summaries,effects=effects,interaction=interaction,holdout=holdout_rows,
                holdout_summary=holdout_summary,selected=selected,
                budget=calibration_budget,verification_budget=verification_budget,
                selection_reason=selection_reason,
                seeds=list(seeds),holdout_seeds=list(holdout))


# =========== EXTENSIÓN PUNTUAL: ESTUDIO DEL CIERRE DE CLÚSTERES ===========
# Estas semillas se fijan ANTES de evaluar tratamientos, y no se seleccionan
# para lograr un costo específico. El grupo de validación es independiente.
SPLIT_EXTENSION_LEVELS = (0.45, 0.50, 0.55, 0.60, 0.65)
SPLIT_EXTENSION_TRAIN_SEEDS = (8301, 8302, 8303, 8304, 8305)
SPLIT_EXTENSION_HOLDOUT_SEEDS = (9301, 9302, 9303, 9304, 9305)


def _split_experiment_run(ins, base, seed, split_ratio, budget=72):
    """Una repetición completa con presupuesto fijo, heurística híbrida intacta."""
    cfg = _replace(base, execution_mode='budget', geometric_p=0.95,
                   cluster_split_ratio=split_ratio, construction_budget=budget,
                   seed=int(seed))
    result = multi_start(ins, cfg)
    row = _summarize_run(result, seed, budget, 'Calibración del cierre CFRS')
    row.update(cluster_split_ratio=split_ratio, geometric_p=0.95,
               clusters=len(result.best_global.clusters) if result.best_global else None,
               combinaciones_con_solucion=len(result.best_by_config),
               construcciones_fallidas=len(result.attempts)-sum(a.feasible for a in result.attempts))
    return row


def _split_summary(rows, ratio):
    sub = [r for r in rows if r['cluster_split_ratio'] == ratio]
    valid = [r for r in sub if r['costo_total'] is not None]
    costs = [r['costo_total'] for r in valid]
    return dict(cluster_split_ratio=ratio, semillas=len(sub),
                ejecuciones_con_solucion=len(valid),
                construcciones=sum(r['construcciones'] for r in sub),
                factibles=sum(r['factibles'] for r in sub),
                costo_medio=_mean(costs) if costs else None,
                costo_min=min(costs) if costs else None,
                costo_sd=_stdev(costs) if len(costs)>1 else (0.0 if costs else None),
                rutas_medias=_mean(r['rutas'] for r in valid) if valid else None,
                clusters_medios=_mean(r['clusters'] for r in valid) if valid else None,
                segundos_medios=_mean(r['tiempo_seg'] for r in sub))


def _split_ranking(row):
    """Factibilidad > costo medio > dispersión > rutas > tiempo, sin pesos."""
    return (-row['ejecuciones_con_solucion'], -row['factibles'],
            row['costo_medio'] if row['costo_medio'] is not None else math.inf,
            row['costo_sd'] if row['costo_sd'] is not None else math.inf,
            row['rutas_medias'] if row['rutas_medias'] is not None else math.inf,
            row['segundos_medios'], -row['cluster_split_ratio'])


def _paired_confidence(train, ratio_a, ratio_b, confidence=0.95):
    """IC t pareado de ahorro (costo de B - costo de A); sin umbral del 1%."""
    from scipy.stats import t as student_t
    a={r['semilla']:r['costo_total'] for r in train if r['cluster_split_ratio']==ratio_a}
    b={r['semilla']:r['costo_total'] for r in train if r['cluster_split_ratio']==ratio_b}
    differences=[b[k]-a[k] for k in sorted(set(a)&set(b))
                 if a[k] is not None and b[k] is not None]
    if len(differences)<2:
        return dict(muestras=len(differences),media=_mean(differences) if differences else None,
                    limite_inferior=None,limite_superior=None)
    avg=_mean(differences)
    width=float(student_t.ppf((1+confidence)/2,len(differences)-1))*_stdev(differences)/math.sqrt(len(differences))
    return dict(muestras=len(differences),media=avg,
                limite_inferior=avg-width,limite_superior=avg+width)


def _controlled_deterministic_comparison(ins, base, seeds, ratio, budget=72):
    """Control determinista PAREADO con idéntica agenda, no una búsqueda distinta.

    Para cada semilla, la búsqueda híbrida establece la secuencia de combinaciones
    y semillas internas. Se reproduce dicha secuencia con p=1 en todas las
    construcciones. Se reportan explícitamente duplicados deterministas.
    La referencia determinista alternativa del factorial previo permanece aparte.
    """
    paired=[]
    for seed in seeds:
        hybrid_config=_replace(base, seed=int(seed), execution_mode='budget',
                               construction_budget=budget, geometric_p=0.95,
                               cluster_split_ratio=ratio)
        hybrid=multi_start(ins,hybrid_config)
        schedule=[a.config_id for a in hybrid.attempts]
        if len(schedule)!=budget:
            raise AssertionError('No hay construcciones suficientes para el control pareado')
        det_cfg=_replace(hybrid_config,geometric_p=1.0)
        deterministic=multi_start(ins,det_cfg,planned_schedule=schedule)
        if [(a.config_id,a.seed) for a in hybrid.attempts] != [
                (a.config_id,a.seed) for a in deterministic.attempts]:
            raise AssertionError('El control pareado cambió la agenda o las semillas internas')
        for label,run in [('HIBRIDA_0.95',hybrid),('DETERMINISTA_AGENDA_IDÉNTICA',deterministic)]:
            row=_summarize_run(run,seed,budget,'Control pareado de selección')
            row.update(metodo=label,cluster_split_ratio=ratio,
                       clusters=len(run.best_global.clusters) if run.best_global else None,
                       agenda='Iguales configuraciones e idénticas semillas de construcción')
            paired.append(row)
        print('CONTROL PAR p=.95/p=1, semilla',seed,'híbrido',
              hybrid.best_global.total_cost if hybrid.best_global else None,
              'determinista',deterministic.best_global.total_cost if deterministic.best_global else None,
              flush=True)
    return paired


def calibrate_cluster_split(ins, base,
              levels=SPLIT_EXTENSION_LEVELS,
              calibration_seeds=SPLIT_EXTENSION_TRAIN_SEEDS,
              holdout_seeds=SPLIT_EXTENSION_HOLDOUT_SEEDS,
              budget=72, deterministic_control=True):
    """Complemento de la calibración previa, NO repite el factorial 4 x 3.

    Cinco cierres con p=0.95 y 72 construcciones híbridas por semilla. Primero
    se selecciona un candidato sobre semillas de calibración. Luego se compara
    con 0.65 en semillas reservadas. Se cambia 0.65 solo si el IC pareado del
    95% de ahorro en holdout es estrictamente positivo, sin pérdida de
    factibilidad. La muestra pequeña limita la generalización de los resultados.
    """
    if budget != 72:
        raise ValueError('Esta ampliación debe comparar exactamente 72 construcciones')
    levels=tuple(levels)
    if 0.65 not in levels or len(set(levels))!=len(levels):
        raise ValueError('Incluir 0.65 como referencia y niveles sin duplicación')
    if set(calibration_seeds)&set(holdout_seeds):
        raise ValueError('Semillas de calibración y validación deben ser disjuntas')
    rows=[]
    for ratio in levels:
        for seed in calibration_seeds:
            r=_split_experiment_run(ins,base,seed,ratio,budget)
            rows.append(r)
        summary=_split_summary(rows,ratio)
        print('CIERRE CALIBRACIÓN',ratio,'medio',summary['costo_medio'],
              'min',summary['costo_min'],'factibles',summary['factibles'],
              '/',summary['construcciones'],'rutas',summary['rutas_medias'],flush=True)
    summary=[_split_summary(rows,r) for r in levels]
    best=min(summary,key=_split_ranking)
    training_candidate=best['cluster_split_ratio']
    finalist=training_candidate
    # El baseline 0.65 se verifica siempre. Si ya ganó, también contrastar
    # el mejor valor más pequeño; nunca elegirlo solo por un mínimo aislado.
    candidates=[x for x in sorted(summary,key=_split_ranking)
                if x['cluster_split_ratio']<0.65]
    if finalist==0.65 and candidates:
        finalist=candidates[0]['cluster_split_ratio']
    holdout=[]
    for ratio in (0.65,finalist) if finalist != 0.65 else (0.65,):
        for seed in holdout_seeds:
            holdout.append(_split_experiment_run(ins,base,seed,ratio,budget))
        summ=_split_summary(holdout,ratio)
        print('CIERRE HOLDOUT',ratio,'costo medio',summ['costo_medio'],
              'rutas',summ['rutas_medias'],'factibles',summ['factibles'],flush=True)
    holdout_summary=[_split_summary(holdout,r) for r in
                     ((0.65,finalist) if finalist!=0.65 else (0.65,))]
    pair=_paired_confidence(holdout,finalist,0.65) if finalist != 0.65 else None
    old=next(x for x in holdout_summary if x['cluster_split_ratio']==0.65)
    new=next((x for x in holdout_summary if x['cluster_split_ratio']==finalist),old)
    accepted=(training_candidate==finalist and finalist!=0.65
              and pair['limite_inferior'] is not None
              and pair['limite_inferior']>0
              and new['ejecuciones_con_solucion']>=old['ejecuciones_con_solucion']
              and new['factibles']>=old['factibles'])
    selected=finalist if accepted else 0.65
    reason=('Menor costo medio confirmado con IC pareado positivo sin perder factibilidad'
            if accepted else 'No se comprobó una mejora consistente sobre 0.65; se conserva el valor previo')
    controlled=(_controlled_deterministic_comparison(ins,base,holdout_seeds,selected,budget)
                if deterministic_control else [])
    print('CIERRE SELECCIONADO:',selected,reason,'IC pareado',pair,flush=True)
    return dict(rows=rows,summary=summary,holdout=holdout,
                holdout_summary=holdout_summary,control=controlled,
                selected=selected,reason=reason,paired_ci=pair,
                budget=budget,calibration_seeds=list(calibration_seeds),
                holdout_seeds=list(holdout_seeds))


def repeatability_test(ins,base,seed=4105,budget=72):
    cfg=_replace(base,seed=seed,construction_budget=budget)
    x=multi_start(ins,cfg);y=multi_start(ins,cfg)
    a,b=_structural_fingerprint(x),_structural_fingerprint(y)
    if a!=b: raise AssertionError('Misma semilla no produjo las mismas soluciones')
    z=multi_start(ins,_replace(cfg,seed=seed+1))
    other=_structural_fingerprint(z)
    if a==other: print('ADVERTENCIA: dos semillas dieron misma huella; revisar aleatorización')
    print(f'REPRODUCIBILIDAD: misma semilla => idéntica {a[:18]}, '
          f'otra semilla => {other[:18]}',flush=True)
    return dict(semilla=seed,construcciones=budget,igual=(a==b),
                distinta=(a!=other),huella=a,huella_otra=other)


# ============== EJECUCIÓN PRINCIPAL CON SEMILLAS Y CIERRES ==============
# Estos niveles son fijos ANTES de ejecutar. Las semillas nunca se eligen
# por resultados previos; se pueden ampliar con --num-seeds o --seeds.
MAIN_SPLIT_LEVELS = (0.45, 0.50, 0.55, 0.60, 0.65)
DEFAULT_MAIN_SEED_COUNT = 3


def run_multiseed_multiratio(ins: Instance, base: SearchConfig,
                             seeds: tuple[int, ...] | list[int],
                             split_levels=MAIN_SPLIT_LEVELS, on_feasible=None):
    """Ejecutar 72 construcciones independientes por (semilla, cierre).

    No mezcla resultados en las memorias de multi_start ni reutiliza ninguna
    solución guardada. Selecciona el mínimo entre soluciones REVALIDADAS con
    costos reales de apertura + transporte. Las métricas de cada ejecución y
    de cada intento se guardan separadamente de la solución principal.
    """
    seeds = tuple(int(v) for v in seeds)
    ratios = tuple(float(v) for v in split_levels)
    if not seeds or len(seeds) != len(set(seeds)):
        raise ValueError('Se necesita al menos una semilla; no se permiten repeticiones')
    if not ratios or len(ratios) != len(set(ratios)) or any(not 0 < v <= 1 for v in ratios):
        raise ValueError('Valores cluster_split_ratio inválidos o repetidos')
    if base.construction_budget != 72:
        raise ValueError('La búsqueda multi-semilla debe usar 72 construcciones por ejecución')
    started = perf_counter()
    summary: list[dict] = []
    attempt_rows: list[dict] = []
    selected_run: RunResult | None = None
    winner_metadata: dict | None = None
    selected_key = None
    total_attempts = 0
    total_feasible = 0

    for index, (outer_seed, ratio) in enumerate(
            ((seed, ratio) for seed in seeds for ratio in ratios), start=1):
        cfg = replace(base, execution_mode='budget', geometric_p=0.95,
                      cluster_split_ratio=ratio, seed=outer_seed,
                      construction_budget=72)
        run = multi_start(ins, cfg, on_feasible=on_feasible)
        count = len(run.attempts)
        feasible_count = sum(a.feasible for a in run.attempts)
        total_attempts += count
        total_feasible += feasible_count
        incumbent = run.best_global
        if incumbent is not None:
            failures = validate_solution(ins, incumbent)
            if not incumbent.feasible or failures:
                raise ValueError('Mejor solución de la ejecución %s no es válida: %s'
                                 % (index, failures))
        metadata = dict(ejecucion=index, semilla_principal=outer_seed,
                        geometric_p=0.95, cluster_split_ratio=ratio,
                        construcciones_solicitadas=72, construcciones_realizadas=count,
                        construcciones_factibles=feasible_count,
                        configuraciones_exploradas=sum(c.attempts > 0 for c in run.combinations),
                        tiempo_busqueda_seg=run.elapsed_seconds,
                        costo_total=incumbent.total_cost if incumbent else None,
                        costo_apertura=incumbent.opening_cost if incumbent else None,
                        costo_transporte=incumbent.transport_cost if incumbent else None,
                        distancia_km=incumbent.distance_km if incumbent else None,
                        rutas=len(incumbent.routes) if incumbent else None,
                        bodegas=' | '.join(incumbent.open_depots) if incumbent else None,
                        configuracion_ganadora=incumbent.config_id if incumbent else None,
                        arranque_ganador=incumbent.start_number if incumbent else None,
                        semilla_interna_ganadora=str(incumbent.seed) if incumbent else None)
        summary.append(metadata)
        for log in run.attempts:
            attempt_rows.append(dict(ejecucion=index,semilla_principal=outer_seed,
                geometric_p=0.95,cluster_split_ratio=ratio,
                solucion_id_local=log.solution_id,config_id=log.config_id,
                arranque=log.start_number,semilla_interna=str(log.seed),
                fase=log.phase,factible=log.feasible,costo_total=log.total_cost if log.feasible else None,
                costo_apertura=log.opening_cost,costo_transporte=log.transport_cost if log.feasible else None,
                rutas=log.routes if log.feasible else None,segundos=log.elapsed_seconds,
                motivo=log.reason))
        # No usar ninguna semilla como desempate de calidad. Si los costos
        # empatan, se comparan distancia, rutas y estructura de modo estable.
        if incumbent is not None:
            key = (incumbent.total_cost, incumbent.distance_km,
                   len(incumbent.routes), signature(incumbent))
            if selected_key is None or key < selected_key:
                selected_key = key
                selected_run = copy.deepcopy(run)
                winner_metadata = dict(metadata)
        # Solo mostrar avance; no presentar los mejores de cada repetición
        # como soluciones finales ni darles prioridad por orden de aparición.
        print('Construcción {}/{} | semilla={} | cierre={:.2f} | '
              'factibles={}/{}'.format(index,len(seeds)*len(ratios),
              outer_seed,ratio,feasible_count,count),flush=True)

    if selected_run is None or winner_metadata is None:
        raise RuntimeError('Ninguna de las ejecuciones encontró una solución factible validada')
    selected_run.settings.update({
        'tipo_ejecucion_principal':'multi-semilla y cinco cierres',
        'semillas_principales_evaluadas':', '.join(map(str,seeds)),
        'cierres_evaluados':', '.join(map(str,ratios)),
        'numero_ejecuciones_independientes':len(summary),
        'construcciones_por_ejecucion':72,
        'construcciones_totales_multi':total_attempts,
        'construcciones_factibles_multi':total_feasible,
        'tiempo_total_busqueda_multi_seg':perf_counter()-started,
        'semilla_principal_elegida':winner_metadata['semilla_principal'],
        'cluster_split_ratio_elegido':winner_metadata['cluster_split_ratio'],
        'geometric_p_elegido':0.95,
        'configuracion_solucion_elegida':winner_metadata['configuracion_ganadora'],
        'arranque_solucion_elegida':winner_metadata['arranque_ganador'],
        'semilla_interna_solucion_elegida':winner_metadata['semilla_interna_ganadora'],
        'criterio_eleccion':'menor costo total real validado entre todas las ejecuciones',
    })
    # La solución final es una sola. El detalle de las demás ejecuciones son
    # MÉTRICAS, no soluciones adicionales ni referencias en la memoria élite.
    return selected_run, dict(resumen=summary,arranques=attempt_rows,
                              mejor=winner_metadata,semillas=list(seeds),
                              cierres=list(ratios),construcciones_totales=total_attempts,
                              factibles_totales=total_feasible,
                              tiempo_total_seg=selected_run.settings['tiempo_total_busqueda_multi_seg'])


def export_integrated(ins,result,output,comparison=None,doe=None,repeatability=None,
                      input_cfg=None,split_study=None,retained_split_sheets=None,
                      multi_search=None):
    # Las hojas de un estudio previo se preservan si las cuatro entradas
    # coinciden SHA-256. El estudio de esta ejecución prevalece sobre ellas.
    extras=dict(retained_split_sheets or {})
    # Los experimentos solo aportan hojas de métricas y calibración.
    # Nunca se incorporan sus soluciones a las alternativas ni a los
    # mejores resultados de las 72 construcciones de la ejecución principal.
    # La huella de datos permite verificar que las repeticiones usaron los
    # mismos bytes de los archivos originales (no solamente iguales tamaños).
    hashes=[]
    file_sources = (dict(bodegas=input_cfg.depots_file, clientes=input_cfg.clients_file,
                        distancias=input_cfg.distance_file, tiempos=input_cfg.time_file)
                    if input_cfg is not None else FILES_DEFAULT)
    for label,path in file_sources.items():
        try:
            real=_resolve_file(path)
            sha=_hashlib.sha256(_Path(real).read_bytes()).hexdigest()
            hashes.append(dict(archivo=label,ruta=real,sha256=sha))
        except FileNotFoundError:
            pass
    if hashes: extras['Control_Entradas']=hashes
    if comparison:
        extras.update(Exper_Presupuestos=comparison['rows'],
                      Exper_Pareado=comparison['paired'],
                      Exper_Resumen=comparison['summary'],
                      Exper_Decision=[dict(presupuesto_elegido=comparison['choice'],
                                           justificacion=comparison['reason'])])
    if doe:
        extras.update(DOE_Hibrido_Corridas=doe['rows'],
                      DOE_Hibrido_Resumen=doe['summary'],
                      DOE_Efectos=doe.get('effects',[]),
                      DOE_Interaccion=doe.get('interaction',[]),
                      DOE_Holdout=doe['holdout'],
                      DOE_Holdout_Resumen=doe['holdout_summary'],
                      DOE_Seleccion=[dict(sesgo=doe['selected'][0],
                            cierre=doe['selected'][1],presupuesto_calibracion=doe['budget'],
                            presupuesto_verificacion=doe['verification_budget'],
                            nota='Selección sobre semillas de calibración; holdout no ajusta ganador',
                            justificacion=doe.get('selection_reason'))])
    if split_study:
        extras.update(Cierre_Ensayos=split_study['rows'],
                      Cierre_Resumen=split_study['summary'],
                      Cierre_Holdout=split_study['holdout'],
                      Cierre_Hold_Resumen=split_study['holdout_summary'],
                      Cierre_ControlDet=split_study['control'],
                      Cierre_Seleccion=[dict(cierre_elegido=split_study['selected'],
                          sesgo_geometrico=0.95, presupuesto=split_study['budget'],
                          semillas_calibracion=', '.join(map(str,split_study['calibration_seeds'])),
                          semillas_holdout=', '.join(map(str,split_study['holdout_seeds'])),
                          justificacion=split_study['reason'],
                          ahorro_pareado_ic=repr(split_study['paired_ci']))])
    if repeatability: extras['Control_Repro']=[repeatability]
    if multi_search is not None:
        extras['Busqueda_Multiple'] = multi_search['resumen']
        extras['Arranques_Multiples'] = multi_search['arranques']
        extras['Seleccion_Global'] = [multi_search['mejor']]
    # Vista SOLO PARA EXPORTACIÓN: una única solución completa, la mejor
    # factible hallada en ESTA ejecución. No alterar la memoria del algoritmo,
    # los registros de los 72 intentos ni los resultados de los experimentos.
    # La hoja Alternativas conserva una sola fila con la mejor global para
    # mantener el esquema de importación compatible con la Fase II.
    best = copy.deepcopy(result.best_global)
    export_view = _replace(
        result,
        best_global=best,
        best_by_config={best.config_id: copy.deepcopy(best)} if best else {},
        alternatives=[copy.deepcopy(best)] if best else [],
    )
    return export_results_excel(output,ins,export_view,extras)



# ========== PERSISTENCIA SEGURA DE CALIBRACIÓN POR INSTANCIA ============
# Solo se recupera la selección si los CUATRO Excel coinciden byte a byte
# con los utilizados en el experimento; cambiar los datos invalida la selección.
_SPLIT_SHEETS = (
    'Cierre_Ensayos', 'Cierre_Resumen', 'Cierre_Holdout',
    'Cierre_Hold_Resumen', 'Cierre_ControlDet', 'Cierre_Seleccion',
)


def _source_file_hashes(input_cfg):
    paths = dict(bodegas=input_cfg.depots_file,
                 clientes=input_cfg.clients_file,
                 distancias=input_cfg.distance_file,
                 tiempos=input_cfg.time_file)
    return {key: _hashlib.sha256(_Path(path).read_bytes()).hexdigest()
            for key, path in paths.items()}


def load_saved_split_calibration(path, input_cfg):
    """Leer selección y evidencia desde el Excel previo; no alterar los datos.

    Retorna (datos, motivo). Nunca aplica una selección sin trazabilidad ni
    si los SHA-256 de las cuatro entradas difieren. El usuario puede guardar
    el mismo libro bajo otra ruta con --output o --calibration-excel.
    """
    location = _Path(path).expanduser()
    if not location.is_file():
        return None, 'No existe Excel de calibración guardada'
    try:
        with pd.ExcelFile(location, engine='openpyxl') as workbook:
            sheets = set(workbook.sheet_names)
            if not {'Cierre_Seleccion', 'Control_Entradas'} <= sheets:
                return None, 'El Excel no contiene Cierre_Seleccion y Control_Entradas'
            control = pd.read_excel(workbook, sheet_name='Control_Entradas', dtype=str)
            if not {'archivo', 'sha256'} <= set(control.columns):
                return None, 'La hoja Control_Entradas no tiene las columnas requeridas'
            recorded = {}
            for _, row in control.iterrows():
                name, signature = str(row['archivo']).strip(), str(row['sha256']).strip().lower()
                if name in recorded:
                    return None, 'Huella de entrada duplicada en Control_Entradas'
                recorded[name] = signature
            expected = _source_file_hashes(input_cfg)
            if set(recorded) != set(expected) or any(recorded[k] != v for k,v in expected.items()):
                return None, 'Las huellas SHA-256 no coinciden: calibración de otros datos'
            selection = pd.read_excel(workbook, sheet_name='Cierre_Seleccion')
            if len(selection) != 1:
                return None, 'La selección guardada debe contener exactamente una fila'
            chosen = selection.iloc[0]
            ratio = float(chosen['cierre_elegido'])
            geometric_p = float(chosen['sesgo_geometrico'])
            budget = int(chosen['presupuesto'])
            if (not math.isfinite(ratio) or not 0 < ratio <= 1
                    or not math.isclose(geometric_p, 0.95, abs_tol=1e-10)
                    or budget != 72):
                return None, 'Parámetros o presupuesto no compatibles con el estudio de cierre'
            # Conservar todo el resultado de experimentos entre ejecuciones normales.
            saved_sheets = {}
            for name in _SPLIT_SHEETS:
                if name in sheets:
                    df = pd.read_excel(workbook, sheet_name=name)
                    saved_sheets[name] = df.where(pd.notna(df), None).to_dict('records')
            if not all(k in saved_sheets for k in _SPLIT_SHEETS):
                return None, 'El Excel no conserva todas las hojas del estudio de cierre'
            return dict(ratio=ratio, geometric_p=geometric_p,
                        budget=budget, experiment_sheets=saved_sheets,
                        source=str(location.resolve())), 'Calibración guardada y entradas SHA-256 verificadas'
    except (OSError, ValueError, TypeError, KeyError, ImportError) as exc:
        return None, f'No se pudo validar la calibración guardada: {type(exc).__name__}: {exc}'



# FASE II — únicamente búsqueda local; no se llama a multi_start ni construct_solution
# ============================================================================
from dataclasses import dataclass as _dataclass_ii
from itertools import combinations as _combinations_ii, product as _product_ii
from time import perf_counter as _perf_counter_ii


@_dataclass_ii
class LocalConfig:
    """Límites de esfuerzo (NO alteran demanda, capacidades ni costes)."""
    max_passes: int = 4
    max_evaluations: int = 30000
    max_accepted: int = 150
    # 2 candidatos económicos por operación al cambiar de depósito (no es MILP)
    candidate_insertions: int = 4
    improvement_eps: float = 1e-6
    max_stage_evaluations: int = 7000
    max_stage_accepted: int = 30

    def validate(self):
        if min(self.max_passes, self.max_evaluations, self.max_accepted,
               self.candidate_insertions, self.max_stage_evaluations,
               self.max_stage_accepted) < 1 or self.improvement_eps <= 0:
            raise ValueError('Límites de Fase II deben ser positivos')


@_dataclass_ii
class LocalLog:
    iteration: int
    movement: str
    detail: str
    before_cost: float
    after_cost: float
    saving: float
    affected_depots: str
    routes: int


class LocalOptimizer:
    """First improvement con cuatro vecindarios simples, en orden fijo.

    Todos los ahorros son coste REAL de la solución completa. Para evaluar un
    candidato se reconstruyen los clústeres/segmentos editados y TODOS los
    emparejamientos y rutas de las bodegas tocadas. Esto es necesario para
    garantizar que la asignación húngara siga siendo de coste mínimo. Las demás
    bodegas y sus rutas permanecen inalteradas. No hay aperturas ni cierres.
    """
    def __init__(self, ins: Instance, initial: Solution, config: LocalConfig):
        config.validate()
        issues = validate_solution(ins, initial)
        if issues:
            raise ValueError('Solución inicial de Fase I no válida: ' + '; '.join(issues[:12]))
        self.ins = ins
        self.cfg = config
        self.initial = copy.deepcopy(initial)
        self.current = copy.deepcopy(initial)
        self.count = 0
        self.history: list[LocalLog] = []
        self.started = _perf_counter_ii()
        self.exhausted = False
        self._serial = 0
        self.pass_number = 0
        self.tried = dict(two_opt=0, relocate=0, swap=0, reassignment=0)
        self.accepted_by_stage = dict(two_opt=0, relocate=0, swap=0, reassignment=0)
        self.stage_exhausted = set()

    def _in_budget(self):
        if self.count >= self.cfg.max_evaluations or len(self.history) >= self.cfg.max_accepted:
            self.exhausted = True
            return False
        return True

    def _fresh_id(self, sol: Solution):
        used = {x.id for x in sol.clusters} | {x.id for x in sol.segments}
        i = 1
        while f'CL-{i:05d}' in used or f'SG-{i:05d}' in used:
            i += 1
        return f'CL-{i:05d}', f'SG-{i:05d}'

    def _segment(self, sol: Solution, sid: str) -> Segment:
        return next(x for x in sol.segments if x.id == sid)

    def _update_segment(self, sol: Solution, sid: str) -> None:
        seg = self._segment(sol, sid)
        if not seg.clients:
            raise AssertionError('No puede haber segmento abierto vacío al reconstruir')
        seg.start, seg.end = seg.clients[0], seg.clients[-1]
        seg.internal_km = sum(self.ins.d(a,b) for a,b in zip(seg.clients,seg.clients[1:]))
        seg.demand = sum(self._qty(c,seg.kind) for c in seg.clients)
        cl = next(x for x in sol.clusters if x.id == seg.cluster_id)
        cl.clients = seg.clients.copy()
        cl.demand = seg.demand

    def _qty(self, client: str, kind: str) -> float:
        c = self.ins.clients[client]
        return c.delivery if kind == 'DELIVERY' else c.pickup

    def _remove(self, sol: Solution, client: str, depot: str, kind: str) -> None:
        segs = [s for s in sol.segments if s.depot_id == depot
                and s.kind == kind and client in s.clients]
        if len(segs) != 1:
            raise AssertionError(f'Operación no aparece una sola vez: {depot} {client} {kind}')
        seg = segs[0]
        seg.clients.remove(client)
        if seg.clients:
            self._update_segment(sol, seg.id)
        else:
            sol.clusters = [c for c in sol.clusters if c.id != seg.cluster_id]
            sol.segments = [s for s in sol.segments if s.id != seg.id]

    def _insert(self, sol: Solution, client: str, depot: str, kind: str,
                sid: str | None, index: int = 0) -> None:
        amount = self._qty(client,kind)
        if amount <= EPS:
            return
        if sid is None:
            cid, segid = self._fresh_id(sol)
            sol.clusters.append(Cluster(cid,depot,kind,[client],amount))
            sol.segments.append(Segment(segid,cid,depot,kind,[client],amount,
                                        client,client,0.0))
            return
        seg = self._segment(sol,sid)
        if (seg.depot_id != depot or seg.kind != kind or
                seg.demand + amount > self.ins.vehicle_capacity + EPS):
            raise ConstructionError('Inserción supera capacidad o segmento incorrecto')
        seg.clients.insert(index, client)
        self._update_segment(sol,sid)

    def _fork_affected(self, affected: set[str], source: Solution | None = None) -> Solution:
        """Copia aislada de escritura: clústeres/segmentos SOLO de bodegas afectadas.

        Rutas y emparejamientos no afectados son inmutables durante la prueba;
        _rebuild() sustituye los de depósitos afectados por objetos nuevos.
        Nunca se alteran las estructuras originales hasta aceptar una mejora.
        """
        base = self.current if source is None else source
        trial = copy.copy(base)
        trial.open_depots = base.open_depots.copy()
        trial.assignment = base.assignment.copy()
        trial.clusters = [copy.deepcopy(c) if c.depot_id in affected else c
                          for c in base.clusters]
        trial.segments = [copy.deepcopy(g) if g.depot_id in affected else g
                          for g in base.segments]
        trial.pairings = base.pairings.copy()
        trial.routes = base.routes.copy()
        trial.violations = base.violations.copy()
        return trial

    def _rebuild(self, sol: Solution, affected: set[str]) -> None:
        """Recomponer solo las bodegas tocadas; actualizar costos incrementalmente.

        El resto de rutas y segmentos no cambia: sus importes se toman de la
        solución actual validada. La validación de cada mejora aceptada vuelve
        a calcular todos los arcos, cargas y costos sin depender de esta suma.
        """
        before_routes = [r for r in self.current.routes if r.depot_id in affected]
        old_km = sum(r.km for r in before_routes)
        old_transport = sum(r.cost for r in before_routes)
        sol.pairings = [p for p in sol.pairings if p.depot_id not in affected]
        sol.routes = [r for r in sol.routes if r.depot_id not in affected]
        new_km = 0.0
        new_transport = 0.0
        for depot in sorted(affected):
            segs = [s for s in sol.segments if s.depot_id == depot]
            pairs = match_segments(self.ins, depot, segs)
            rebuilt = build_routes(self.ins, pairs, segs)
            for i, r in enumerate(rebuilt, 1):
                r.id = f'R-{depot}-{i:04d}'
                new_km += r.km
                new_transport += r.cost
            sol.pairings.extend(pairs)
            sol.routes.extend(rebuilt)
        sol.clusters.sort(key=lambda c:(c.depot_id,c.kind,c.id))
        sol.segments.sort(key=lambda s:(s.depot_id,s.kind,s.id))
        sol.pairings.sort(key=lambda p:(p.depot_id,p.delivery_segment or '',p.pickup_segment or ''))
        sol.routes.sort(key=lambda r:(r.depot_id,r.id))
        old_open = set(self.current.open_depots)
        new_open = set(sol.open_depots)
        sol.opening_cost = (self.current.opening_cost +
            sum(self.ins.depots[x].opening_cost for x in new_open-old_open) -
            sum(self.ins.depots[x].opening_cost for x in old_open-new_open))
        sol.transport_cost = self.current.transport_cost-old_transport+new_transport
        sol.distance_km = self.current.distance_km-old_km+new_km
        sol.vehicle_rental_cost = len(sol.routes)*self.ins.vehicle_rental_daily
        sol.total_cost = sol.opening_cost + sol.vehicle_rental_cost + sol.transport_cost

    def _evaluate(self, candidate: Solution, affected: set[str],
                  movement: str, detail: str) -> bool:
        if not self._in_budget():
            return False
        if (self.tried[movement] >= self.cfg.max_stage_evaluations or
                self.accepted_by_stage[movement] >= self.cfg.max_stage_accepted):
            self.stage_exhausted.add(movement)
            return False
        self.count += 1
        self.tried[movement] += 1
        self._rebuild(candidate,affected)
        if candidate.total_cost >= self.current.total_cost - self.cfg.improvement_eps:
            return False
        violations = validate_solution(self.ins,candidate)
        if violations:
            # No se ocultan fallos de código como 'sin mejora'.
            raise AssertionError('Movimiento de Fase II generó solución inválida: '
                                 + '; '.join(violations[:12]) + f' | {movement}: {detail}')
        old = self.current.total_cost
        self.history.append(LocalLog(self.pass_number,movement,detail,old,
                     candidate.total_cost,old-candidate.total_cost,','.join(sorted(affected)),
                     len(candidate.routes)))
        self.accepted_by_stage[movement] += 1
        self.current = copy.deepcopy(candidate)
        return True

    def _edit_trial(self, seg_changes: dict[str,list[str]], affected: set[str],
                    movement: str, detail: str,
                    transfer: dict[str,str] | None = None) -> bool:
        # Descartar ediciones que reproducen exactamente el estado actual.
        if (not transfer and all(self._segment(self.current, sid).clients == seq
                                 for sid,seq in seg_changes.items())):
            return False
        trial = self._fork_affected(affected)
        for sid,seq in seg_changes.items():
            seg = self._segment(trial,sid)
            seg.clients = seq.copy()
            if not seq:
                trial.clusters = [x for x in trial.clusters if x.id != seg.cluster_id]
                trial.segments = [x for x in trial.segments if x.id != sid]
            else:
                self._update_segment(trial,sid)
        if transfer:
            trial.assignment.update(transfer)
        return self._evaluate(trial,affected,movement,detail)

    def two_opt(self) -> bool:
        # Dos-opt aplicado a segmentos abiertos, nunca entre tipos de operación.
        for seg in sorted(self.current.segments,key=lambda x:(x.depot_id,x.id)):
            n = len(seg.clients)
            if n < 2:
                continue
            for i in range(n-1):
                for j in range(i+1,n):
                    seq = seg.clients[:i] + list(reversed(seg.clients[i:j+1])) + seg.clients[j+1:]
                    if self._edit_trial({seg.id:seq},{seg.depot_id},'two_opt',
                                        f'{seg.id} invertir [{i}:{j}]'):
                        return True
                    if self.exhausted or 'two_opt' in self.stage_exhausted:
                        return False
        return False

    def relocate(self) -> bool:
        # Operaciones individuales entre dos rutas de una misma bodega.
        for depot in sorted(self.current.open_depots):
            routes = sorted((r for r in self.current.routes if r.depot_id == depot),
                            key=lambda x:x.id)
            for src in routes:
                for kind in ('DELIVERY','PICKUP'):
                    src_sid = src.delivery_segment if kind=='DELIVERY' else src.pickup_segment
                    if not src_sid:
                        continue
                    source = self._segment(self.current,src_sid)
                    for dst in routes:
                        if src.id == dst.id:
                            continue
                        dst_sid = dst.delivery_segment if kind=='DELIVERY' else dst.pickup_segment
                        target = self._segment(self.current,dst_sid) if dst_sid else None
                        for client in source.clients:
                            amount = self._qty(client,kind)
                            if target and target.demand + amount > self.ins.vehicle_capacity + EPS:
                                continue
                            positions = range(len(target.clients)+1) if target else (0,)
                            for pos in positions:
                                trial = self._fork_affected({depot})
                                self._remove(trial,client,depot,kind)
                                self._insert(trial,client,depot,kind,dst_sid,pos)
                                if self._evaluate(trial,{depot},'relocate',
                                  f'{client} {kind} {src.id}->{dst.id} pos={pos}'):
                                    return True
                                if self.exhausted or 'relocate' in self.stage_exhausted: return False
        return False

    def swap(self) -> bool:
        # Intercambia operaciones de IGUAL tipo entre rutas de una bodega.
        for depot in sorted(self.current.open_depots):
            routes = sorted((r for r in self.current.routes if r.depot_id == depot),key=lambda r:r.id)
            for r1,r2 in _combinations_ii(routes,2):
                for kind in ('DELIVERY','PICKUP'):
                    a = r1.delivery_segment if kind=='DELIVERY' else r1.pickup_segment
                    b = r2.delivery_segment if kind=='DELIVERY' else r2.pickup_segment
                    if not a or not b or a == b:
                        continue
                    sa, sb = self._segment(self.current,a),self._segment(self.current,b)
                    for i,c1 in enumerate(sa.clients):
                        for j,c2 in enumerate(sb.clients):
                            v1,v2=self._qty(c1,kind),self._qty(c2,kind)
                            if sa.demand-v1+v2 > self.ins.vehicle_capacity+EPS: continue
                            if sb.demand-v2+v1 > self.ins.vehicle_capacity+EPS: continue
                            xa,xb=sa.clients.copy(),sb.clients.copy()
                            xa[i],xb[j] = c2,c1
                            if self._edit_trial({a:xa,b:xb},{depot},'swap',
                                       f'{c1}<->{c2} {kind} {r1.id}/{r2.id}'):
                                return True
                            if self.exhausted or 'swap' in self.stage_exhausted: return False
        return False

    def _choose_insertions(self, sol: Solution, client: str, depot: str,
                           kind: str, max_choices: int) -> list[tuple[float,str|None,int]]:
        amount=self._qty(client,kind)
        if amount <= EPS:
            return [(0.0,None,0)]
        options=[]
        for s in sol.segments:
            if s.depot_id!=depot or s.kind!=kind or s.demand+amount > self.ins.vehicle_capacity+EPS:
                continue
            for pos in range(len(s.clients)+1):
                delta=_insertion_delta(self.ins,depot,kind,s.clients,client,pos,closed=True)
                options.append((float(delta),s.id,pos))
        # Abrir un segmento adicional NO abre una bodega ni crea costos fijos de vehículo.
        solo=self.ins.d(depot,client)+self.ins.d(client,depot)
        options.append((float(solo),None,0))
        options.sort(key=lambda x:(x[0],x[1] or '~',x[2]))
        return options[:max_choices]

    def _transfer_candidate(self, moves: dict[str,str],
                            choice_limit: int) -> list[Solution]:
        """Cada traslado mueve ambas operaciones de un cliente, sin dividirlas.

        Para swaps se extraen ambas personas ANTES de insertar cualquiera;
        evita comprobar capacidades contra demandas que ya salieron.
        """
        affected={self.current.assignment[c] for c in moves} | set(moves.values())
        base=self._fork_affected(affected)
        old_depots={c:base.assignment[c] for c in moves}
        new_loads = {d:sum(self.ins.clients[c].delivery for c,assigned in base.assignment.items()
                           if assigned==d) for d in base.open_depots}
        for c,new in moves.items():
            old=old_depots[c]
            if old==new: return []
            new_loads[old]-=self.ins.clients[c].delivery
            new_loads[new]+=self.ins.clients[c].delivery
        if any(v > self.ins.depots[d].capacity+EPS or v < -EPS
               for d,v in new_loads.items()):
            return []
        for c in sorted(moves):
            old=old_depots[c]
            for kind in ('DELIVERY','PICKUP'):
                if self._qty(c,kind)>EPS:
                    self._remove(base,c,old,kind)
        base.assignment.update(moves)
        ordered=sorted(moves)
        options=[]
        for c in ordered:
            depot=moves[c]
            line=self._choose_insertions(base,c,depot,'DELIVERY',choice_limit)
            back=self._choose_insertions(base,c,depot,'PICKUP',choice_limit)
            options.append([(c,depot,dl,pl,db,pb) for (_,dl,pl),(_,db,pb) in _product_ii(line,back)])
        # Si ambos clientes del swap se mueven en sentidos contrarios, no se
        # insertan en la misma bodega: elegir sus destinos es independiente.
        result=[]
        for plan in _product_ii(*options):
            trial=self._fork_affected(affected, base)
            try:
                for c,depot,dl,pl,db,pb in plan:
                    if self._qty(c,'DELIVERY')>EPS:
                        self._insert(trial,c,depot,'DELIVERY',dl,pl)
                    if self._qty(c,'PICKUP')>EPS:
                        self._insert(trial,c,depot,'PICKUP',db,pb)
            except ConstructionError:
                continue
            result.append(trial)
        return result

    def reassignment(self) -> bool:
        # 1) Cliente completo (entrega + recogida) a otra bodega ABIERTA.
        clients=sorted(self.current.assignment)
        opened=sorted(self.current.open_depots)
        for c in clients:
            current=self.current.assignment[c]
            for depot in opened:
                if depot==current: continue
                for trial in self._transfer_candidate({c:depot},self.cfg.candidate_insertions):
                    if self._evaluate(trial,{current,depot},'reassignment',
                                      f'trasladar cliente {c} {current}->{depot}'):
                        return True
                    if self.exhausted or 'reassignment' in self.stage_exhausted: return False
        # 2) Intercambio de clientes completos de bodegas diferentes.
        for c1,c2 in _combinations_ii(clients,2):
            d1,d2=self.current.assignment[c1],self.current.assignment[c2]
            if d1==d2: continue
            for trial in self._transfer_candidate({c1:d2,c2:d1},1):
                if self._evaluate(trial,{d1,d2},'reassignment',
                        f'intercambiar clientes {c1}/{c2} {d1}<->{d2}'):
                    return True
                if self.exhausted or 'reassignment' in self.stage_exhausted: return False
        return False

    def optimize(self) -> Solution:
        stages=[('2-opt',self.two_opt),('Relocate',self.relocate),
                ('Swap',self.swap),('Reasignación',self.reassignment)]
        for iteration in range(1,self.cfg.max_passes+1):
            self.pass_number=iteration
            changed=False
            for label,method in stages:
                key={'2-opt':'two_opt','Relocate':'relocate','Swap':'swap',
                     'Reasignación':'reassignment'}[label]
                while (self._in_budget() and key not in self.stage_exhausted and method()):
                    changed=True
                    log=self.history[-1]
                    print(f'  {label}: ahorro {log.saving:,.2f} COP, '
                          f'costo {log.after_cost:,.2f} COP | {log.detail}',flush=True)
            if self.exhausted or not changed:
                break
        final=copy.deepcopy(self.current)
        final.id='FASE-II-FINAL'
        final.feasible=True
        final.violations=[]
        violations=validate_solution(self.ins,final)
        if violations:
            raise AssertionError('Solución final inválida: '+str(violations[:12]))
        if final.total_cost > self.initial.total_cost+EPS:
            raise AssertionError('La Fase II no puede empeorar el costo')
        return final



# ====================== INTEGRACIÓN DE FASES ============================
# Las funciones constructivas anteriores son las originales. Únicamente se
# conectan los resultados válidos y se implementa aceptación BEST improvement.

from io import BytesIO
import argparse as _parser_integrado
from datetime import datetime as _datetime_integrado


class CincoMejoresFaseI:
    """Top cinco EXACTO entre todas las construcciones válidas, sin diversidad artificial.

    La firma canónica comprueba instalaciones, asignaciones y rutas. Cada
    registro es una copia profunda independiente; no se retienen soluciones
    experimentales ni se reconstruye la Fase I.
    """
    def __init__(self, ins: Instance, limit: int = 5):
        self.ins = ins
        self.limit = limit
        self.pool: dict[tuple, Solution] = {}
        self.recibidas = 0
        self.distintas = 0

    @staticmethod
    def criterio(sol: Solution):
        return (round(sol.total_cost, 6), round(sol.distance_km, 6),
                len(sol.routes), signature(sol))

    def incluir(self, candidate: Solution):
        if not candidate.feasible or candidate.violations:
            raise AssertionError('Solo se aceptan soluciones de Fase I validadas')
        self.recibidas += 1
        key = signature(candidate)
        old = self.pool.get(key)
        if old is not None:
            if self.criterio(candidate) < self.criterio(old):
                self.pool[key] = copy.deepcopy(candidate)
            return
        self.distintas += 1
        if len(self.pool) < self.limit:
            self.pool[key] = copy.deepcopy(candidate)
            return
        worst = max(self.pool, key=lambda k:self.criterio(self.pool[k]))
        if self.criterio(candidate) < self.criterio(self.pool[worst]):
            del self.pool[worst]
            self.pool[key] = copy.deepcopy(candidate)

    def finales(self):
        selected = sorted(self.pool.values(), key=self.criterio)
        ret = []
        for i, sol in enumerate(selected, 1):
            if validate_solution(self.ins,sol):
                raise AssertionError('Top de Fase I contiene solución inválida')
            replica = copy.deepcopy(sol)
            replica.id = f'FI-ELITE-{i:02d}'
            ret.append(replica)
        return ret


class BusquedaLocalMejorMejora(LocalOptimizer):
    """Mismos cuatro vecindarios, cambiando solamente first por best improvement.

    Cada barrido examina todos los candidatos que permita el límite explícito
    del vecindario, compara el costo REAL de las rutas reemparejadas y aplica
    exactamente el mayor ahorro estricto. Después reinicia ese vecindario.
    Todos los desempates dependen del orden canónico, nunca del reloj.
    """
    def __init__(self, ins, initial, config):
        super().__init__(ins, initial, config)
        # La fusión tiene el mismo presupuesto independiente que los otros movimientos.
        self.tried['fusion'] = 0
        self.accepted_by_stage['fusion'] = 0
        self._mejor_barrido = None
        self._detalle_mejor = None

    def relocate(self) -> bool:
        """Reinsertar dentro del mismo segmento o entre rutas del mismo depósito.

        El movimiento dentro de una ruta nunca altera la operación de tipo opuesto.
        Las otras bodegas permanecen intactas al reconstruir el candidato.
        """
        for depot in sorted(self.current.open_depots):
            routes = sorted((r for r in self.current.routes if r.depot_id == depot),
                            key=lambda r: r.id)
            for src in routes:
                for kind in ('DELIVERY','PICKUP'):
                    sid = src.delivery_segment if kind == 'DELIVERY' else src.pickup_segment
                    if not sid:
                        continue
                    segment = self._segment(self.current,sid)
                    # Inserciones dentro de la misma ruta (incluidas posiciones no vecinas).
                    for i, client in enumerate(segment.clients):
                        rest = segment.clients[:i] + segment.clients[i+1:]
                        positions=self._swap_positions(depot,kind,rest,client)
                        for j in positions:
                            revised = rest[:j] + [client] + rest[j:]
                            if revised == segment.clients:
                                continue
                            self._edit_trial({sid: revised},{depot},'relocate',
                                f'{client} {kind} dentro de {src.id}: {i}->{j}')
                            if self.exhausted or 'relocate' in self.stage_exhausted:
                                return False
                    # Inserciones en otras rutas de la misma bodega.
                    for dst in routes:
                        if dst.id == src.id:
                            continue
                        tid = dst.delivery_segment if kind == 'DELIVERY' else dst.pickup_segment
                        target = self._segment(self.current,tid) if tid else None
                        for client in segment.clients:
                            amount = self._qty(client,kind)
                            if target and target.demand + amount > self.ins.vehicle_capacity + EPS:
                                continue
                            positions = (self._swap_positions(depot,kind,target.clients,client)
                                         if target else (0,))
                            for j in positions:
                                trial = self._fork_affected({depot})
                                self._remove(trial,client,depot,kind)
                                self._insert(trial,client,depot,kind,tid,j)
                                self._evaluate(trial,{depot},'relocate',
                                    f'{client} {kind} {src.id}->{dst.id} pos={j}')
                                if self.exhausted or 'relocate' in self.stage_exhausted:
                                    return False
        return False

    def _swap_positions(self, depot: str, kind: str, remaining: list[str],
                        incoming: str) -> list[int]:
        """Hasta cuatro mejores posiciones según incremento vial DIRIGIDO.

        Es un filtro de candidatos, no el costo final: _evaluate calcula
        el costo real de todas las rutas afectadas y valida la solución.
        """
        choices = [(_insertion_delta(self.ins,depot,kind,remaining,incoming,pos),pos)
                   for pos in range(len(remaining)+1)]
        choices.sort(key=lambda x:(x[0],x[1]))
        return [position for _,position in choices[:self.cfg.candidate_insertions]]

    def swap(self) -> bool:
        """Intercambiar operaciones del mismo tipo y explorar posiciones alternativas."""
        for depot in sorted(self.current.open_depots):
            routes = sorted((r for r in self.current.routes if r.depot_id == depot),
                            key=lambda r:r.id)
            for r1,r2 in _combinations_ii(routes,2):
                for kind in ('DELIVERY','PICKUP'):
                    a = r1.delivery_segment if kind == 'DELIVERY' else r1.pickup_segment
                    b = r2.delivery_segment if kind == 'DELIVERY' else r2.pickup_segment
                    if not a or not b or a == b:
                        continue
                    sa,sb = self._segment(self.current,a),self._segment(self.current,b)
                    for i,c1 in enumerate(sa.clients):
                        for j,c2 in enumerate(sb.clients):
                            v1,v2 = self._qty(c1,kind),self._qty(c2,kind)
                            if sa.demand-v1+v2 > self.ins.vehicle_capacity+EPS:
                                continue
                            if sb.demand-v2+v1 > self.ins.vehicle_capacity+EPS:
                                continue
                            left = sa.clients[:i]+sa.clients[i+1:]
                            right = sb.clients[:j]+sb.clients[j+1:]
                            p_left = self._swap_positions(depot,kind,left,c2)
                            p_right = self._swap_positions(depot,kind,right,c1)
                            for ia in p_left:
                                xa = left[:ia]+[c2]+left[ia:]
                                for ib in p_right:
                                    xb = right[:ib]+[c1]+right[ib:]
                                    self._edit_trial({a:xa,b:xb},{depot},'swap',
                                        f'{c1}<->{c2} {kind} {r1.id}/{r2.id} '
                                        f'pos={ia}/{ib}')
                                    if self.exhausted or 'swap' in self.stage_exhausted:
                                        return False
        return False

    def fusion(self) -> bool:
        """Fusionar rutas del mismo depósito si las demandas combinadas caben.

        Se unen los segmentos de cada tipo y se eliminan los segmentos vacíos.
        Después el emparejamiento húngaro reconstruye solo las rutas del
        depósito afectado. No se exige menos rutas ni se penaliza su número;
        únicamente se conserva una reducción estricta del costo real.
        """
        for depot in sorted(self.current.open_depots):
            routes = sorted((r for r in self.current.routes if r.depot_id == depot),
                            key=lambda r:r.id)
            for r1,r2 in _combinations_ii(routes,2):
                segment_changes = []
                for kind in ('DELIVERY','PICKUP'):
                    a = r1.delivery_segment if kind == 'DELIVERY' else r1.pickup_segment
                    b = r2.delivery_segment if kind == 'DELIVERY' else r2.pickup_segment
                    if a and b and a != b:
                        sa,sb = self._segment(self.current,a),self._segment(self.current,b)
                        if sa.demand+sb.demand > self.ins.vehicle_capacity+EPS:
                            break
                        segment_changes.append((a,b,sa.clients,sb.clients))
                else:
                    # Si no tienen operaciones de un mismo tipo, su unión ya
                    # corresponde a una opción del emparejamiento existente.
                    if not segment_changes:
                        continue
                    alternatives=[]
                    for a,b,seq_a,seq_b in segment_changes:
                        alternatives.append([(a,b,seq_a+seq_b),
                                             (a,b,seq_b+seq_a)])
                    for combined in _product_ii(*alternatives):
                        changes={}
                        for a,b,seq in combined:
                            changes[a]=seq
                            changes[b]=[]
                        self._edit_trial(changes,{depot},'fusion',
                                         f'{r1.id}+{r2.id} '
                                         + '/'.join(x[0] for x in combined))
                        if self.exhausted or 'fusion' in self.stage_exhausted:
                            return False
        return False

    def _evaluate(self, candidate: Solution, affected: set[str],
                  movement: str, detail: str) -> bool:
        # Clave estructural ANTES de rehacer rutas. Evita validar/reemparejar
        # soluciones iguales generadas por rutas/posiciones diferentes.
        touched = tuple(sorted(affected))
        fingerprint = (touched,
            tuple(sorted((x.id, tuple(x.clients)) for x in candidate.segments
                         if x.depot_id in affected)),
            tuple(sorted((c, d) for c, d in candidate.assignment.items()
                         if d in affected or self.current.assignment.get(c) in affected)))
        if fingerprint in self._trial_signatures_seen:
            return False
        self._trial_signatures_seen.add(fingerprint)
        if not self._in_budget():
            return False
        if (self.tried[movement] >= self.cfg.max_stage_evaluations or
            self.accepted_by_stage[movement] >= self.cfg.max_stage_accepted):
            self.stage_exhausted.add(movement)
            return False
        self.count += 1
        self.tried[movement] += 1
        self._rebuild(candidate, affected)
        # Jamás aceptar empates: siempre mejora estricta sobre incumbente.
        if candidate.total_cost >= self.current.total_cost-self.cfg.improvement_eps:
            return False
        if (self._mejor_barrido is not None and
            candidate.total_cost >= self._mejor_barrido.total_cost-self.cfg.improvement_eps):
            return False
        failures = validate_solution(self.ins, candidate)
        if failures:
            raise AssertionError('Candidato con costo menor pero inválido: '
                                 + '; '.join(failures[:12]) + ' | '+movement+' '+detail)
        self._mejor_barrido = candidate
        self._detalle_mejor = (movement, detail, tuple(sorted(affected)))
        # False obliga a los bucles de la clase constructiva a seguir
        # evaluando el resto del vecindario en lugar de parar en el primero.
        return False

    def _aceptar_mejor_barrido(self) -> bool:
        if self._mejor_barrido is None:
            return False
        candidate = self._mejor_barrido
        movement, detail, affected = self._detalle_mejor
        before = self.current.total_cost
        if not before-candidate.total_cost > self.cfg.improvement_eps:
            raise AssertionError('No se aceptan movimientos sin ahorro estricto')
        self.history.append(LocalLog(self.pass_number,movement,detail,before,
                 candidate.total_cost,before-candidate.total_cost,','.join(affected),
                 len(candidate.routes)))
        self.accepted_by_stage[movement] += 1
        self.current = copy.deepcopy(candidate)
        self._mejor_barrido = None
        self._detalle_mejor = None
        return True

    def optimize(self) -> Solution:
        stages = [('2-opt','two_opt',self.two_opt),
                  ('Relocate','relocate',self.relocate),
                  ('Swap','swap',self.swap),
                  ('Fusión','fusion',self.fusion),
                  ('Reasignación','reassignment',self.reassignment)]
        for k in range(1,self.cfg.max_passes+1):
            self.pass_number = k
            any_change = False
            for label,key,search in stages:
                while key not in self.stage_exhausted and self._in_budget():
                    self._mejor_barrido = None
                    self._detalle_mejor = None
                    # Los métodos originales recorren el vecindario en orden
                    # canónico; _evaluate acumula el mejor, no acepta el primero.
                    search()
                    if not self._aceptar_mejor_barrido():
                        break
                    any_change = True
                    change = self.history[-1]
                    print('    {}: ahorro {:,.0f} COP; costo {:,.0f} COP'.format(
                        label,change.saving,change.after_cost),flush=True)
                if self.exhausted:
                    break
            if not any_change or self.exhausted:
                break
        result = copy.deepcopy(self.current)
        result.feasible = True
        result.violations = validate_solution(self.ins,result)
        if result.violations:
            raise AssertionError('Búsqueda local produce solución no factible: '
                                 +'; '.join(result.violations[:10]))
        if result.total_cost > self.initial.total_cost+1e-6:
            raise AssertionError('Fase II empeoró la solución inicial')
        return result


# ====== EXTENSIÓN ACOTADA: APERTURA Y TRASLADO DE CLIENTES COMPLETOS ======
# No modifica la construcción ni los vecindarios originales. Se añade al final.
from itertools import chain as _chain_open, combinations as _combinations_open
from hashlib import sha256 as _sha256_open

@_dataclass_ii(frozen=True)
class OpeningConfig:
    enabled: bool = True
    small_instance_threshold: int = 5
    max_groups_large: int = 28
    max_depots_large: int = 6
    max_evaluations: int = 350  # 0 = sin tope; independiente de otros movimientos
    max_accepted: int = 20     # 0 = sin tope
    max_seconds: float = 0.0   # 0 = sin tope

    def validate(self):
        if (self.small_instance_threshold < 0 or self.max_groups_large < 1 or
            self.max_depots_large < 1 or self.max_evaluations < 0 or
            self.max_accepted < 0 or self.max_seconds < 0):
            raise ValueError('Configuración de apertura inválida')


class BusquedaLocalConApertura(BusquedaLocalMejorMejora):
    """Conserva cinco vecindarios originales, añade apertura como sexto.

    Apertura: examina grupos de clientes completos hacia una bodega CERRADA.
    Para 5 o menos clientes explora todas las combinaciones de clientes y
    destinos cerrados, salvo límite de seguridad explícito. Para mayores,
    prioriza grupos de rutas y proximidad; la exploración CONTROLADA no prueba
    optimalidad local global. Reconstruye con las rutinas originales solamente
    las bodegas de origen y la nueva bodega. Nunca cierra bodegas.
    """
    def __init__(self, ins, initial, config, opening_config=None, construction_config=None):
        super().__init__(ins,initial,config)
        self.open_cfg = opening_config or OpeningConfig()
        self.open_cfg.validate()
        self.construct_cfg = construction_config or SearchConfig()
        self.tried['opening']=0
        self.accepted_by_stage['opening']=0
        self.opening_started = _perf_counter_ii()
        self.opening_incomplete = False
        self.opening_limit_reason = ''
        self.times_by_stage = []
        self.economic_moves = []
        self._opening_candidates_seen = set()
        self._trial_signatures_seen = set()

    def _in_budget(self):
        # Los éxitos de Apertura tienen su PROPIA cuota y no consumen el
        # máximo global de mejoras de los cinco movimientos originales.
        old_accepted=sum(h.movement!='opening' for h in self.history)
        if self.count>=self.cfg.max_evaluations or old_accepted>=self.cfg.max_accepted:
            self.exhausted=True
            return False
        return True

    def _aceptar_mejor_barrido(self):
        before = self.current
        if self._mejor_barrido is None:
            return False
        after = self._mejor_barrido
        old_assign=before.assignment
        new_assign=after.assignment
        changed=[c for c in sorted(old_assign) if old_assign[c]!=new_assign[c]]
        old_open=set(before.open_depots)
        new_open=set(after.open_depots)
        snap = dict(bodegas_antes=' | '.join(sorted(old_open)),
            bodegas_despues=' | '.join(sorted(new_open)),
            bodegas_abiertas=' | '.join(sorted(new_open-old_open)),
            bodegas_cerradas=' | '.join(sorted(old_open-new_open)),
            clientes_trasladados=' | '.join(changed),n_trasladados=len(changed),
            bodegas_costo_antes=before.opening_cost,bodegas_costo_despues=after.opening_cost,
            arriendo_antes=before.vehicle_rental_cost,arriendo_despues=after.vehicle_rental_cost,
            transporte_antes=before.transport_cost,transporte_despues=after.transport_cost,
            km_antes=before.distance_km,km_despues=after.distance_km,
            rutas_antes=len(before.routes),rutas_despues=len(after.routes))
        done=super()._aceptar_mejor_barrido()
        if done:
            # Las mismas estructuras pueden volver a ser prometedoras después
            # de cambiar el incumbente; reiniciar solo la memoria de pruebas.
            self._trial_signatures_seen.clear()
            self._opening_candidates_seen.clear()
            errors=validate_solution(self.ins,self.current)
            if errors:
                raise AssertionError('Solución aceptada no válida: '+'; '.join(errors[:12]))
            if not before.total_cost-self.current.total_cost>self.cfg.improvement_eps:
                raise AssertionError('Un movimiento aceptado no redujo el costo diario')
            log=self.history[-1]
            snap.update(pasada=log.iteration,movimiento=log.movement,detalle=log.detail,
                costo_antes=log.before_cost,costo_despues=log.after_cost,
                ahorro=log.saving,
                ahorro_porcentaje=(100*log.saving/log.before_cost if log.before_cost else 0.0),
                evaluaciones_acumuladas_del_movimiento=self.tried.get(log.movement,0),
                segundos_desde_inicio_optimizador=_perf_counter_ii()-self.started,
                bodegas_afectadas=log.affected_depots)
            self.economic_moves.append(snap)
        return done

    def _candidate_groups(self, target):
        available=sorted(c for c,d in self.current.assignment.items() if d!=target)
        if len(self.ins.clients)<=self.open_cfg.small_instance_threshold:
            return list(_chain_open.from_iterable(
                _combinations_open(available,k) for k in range(1,len(available)+1))),False
        # Grupos asociados a rutas actuales; todos sus clientes se trasladan
        # completos, aunque la entrega y recogida ocupen rutas diferentes.
        proposed=[]
        for route in self.current.routes:
            group=tuple(sorted(set(st.node_id for st in route.stops[1:-1])))
            if group:
                proposed.append(group)
        for d in sorted(self.current.open_depots):
            group=tuple(sorted(c for c,x in self.current.assignment.items() if x==d))
            if group: proposed.append(group)
        # Incluir traslados individuales sin alterar la asignación constructiva.
        proposed.extend([(c,) for c in available])
        by_depot={}
        for route in self.current.routes:
            by_depot.setdefault(route.depot_id,[]).append(
                tuple(sorted(set(st.node_id for st in route.stops[1:-1]))))
        for d,groups in by_depot.items():
            for a,b in _combinations_open(groups,2):
                proposed.append(tuple(sorted(set(a)|set(b))))
        groups=list(dict.fromkeys(g for g in proposed if g and all(self.current.assignment[c]!=target for c in g)))
        # Capacidad es una condición necesaria y se evalúa antes de construir.
        capacity=self.ins.depots[target].capacity
        groups=[g for g in groups if sum(self.ins.clients[c].delivery for c in g)<=capacity+EPS]
        # Ranking aproximado únicamente. NO descartar por costo proxy.
        def rank(group):
            estimate=sum(self.ins.cost_per_km*(
                self.ins.d(target,c)+self.ins.d(c,target)-
                self.ins.d(self.current.assignment[c],c)-
                self.ins.d(c,self.current.assignment[c])) for c in group)
            return (estimate,len(group),group)
        groups.sort(key=rank)
        truncated=len(groups)>self.open_cfg.max_groups_large
        return groups[:self.open_cfg.max_groups_large],truncated

    def _reconstruct_group(self, target, group):
        ins=self.ins
        origins={self.current.assignment[c] for c in group}
        affected=origins|{target}
        trial=self._fork_affected(affected)
        trial.open_depots=sorted(set(trial.open_depots)|{target})
        for c in group:
            trial.assignment[c]=target
        # Descartar antes de reconstruir rutas: capacidad de entrega POR bodega.
        for depot in affected:
            demand=sum(ins.clients[c].delivery for c,x in trial.assignment.items() if x==depot)
            if demand>ins.depots[depot].capacity+EPS:
                return None,affected
        trial.clusters=[cl for cl in trial.clusters if cl.depot_id not in affected]
        trial.segments=[sg for sg in trial.segments if sg.depot_id not in affected]
        # Se reutilizan directamente CFRS y Cheapest Insertion preexistentes.
        next_id=max([int(x.id.split('-')[-1]) for x in trial.clusters
                     if x.id.split('-')[-1].isdigit()] or [0])+1
        for depot in sorted(affected):
            if not any(x==depot for x in trial.assignment.values()):
                continue
            seed_material=f'{depot}|{target}|{";".join(group)}|{self.current.seed}'
            seed=int.from_bytes(_sha256_open(seed_material.encode()).digest()[:8],'big')
            rng=random.Random(seed)
            for kind in ('DELIVERY','PICKUP'):
                clusters=cluster_operations(ins,depot,trial.assignment,kind,
                    self.construct_cfg.geometric_p,rng,next_id,
                    self.construct_cfg.cluster_split_ratio)
                next_id+=len(clusters)
                trial.clusters.extend(clusters)
                trial.segments.extend(cheapest_insertion(ins,x) for x in clusters)
        self._rebuild(trial,affected)
        return trial,affected

    def apertura_localizacion(self):
        if not self.open_cfg.enabled:
            return False
        if self.open_cfg.max_accepted and self.accepted_by_stage['opening']>=self.open_cfg.max_accepted:
            self.stage_exhausted.add('opening')
            self.opening_limit_reason='LIMITE_MEJORAS_APERTURA'
            return False
        closed=sorted(set(self.ins.depots)-set(self.current.open_depots))
        if not closed:
            return False
        small=len(self.ins.clients)<=self.open_cfg.small_instance_threshold
        if not small and len(closed)>self.open_cfg.max_depots_large:
            self.opening_incomplete=True
        # Proximidad estimada de las bodegas actualmente asignadas.
        def target_rank(d):
            # Conservar el ranking original de proximidad. El costo diario
            # se compara de manera exacta después de reconstruir cada opción.
            # Solo en empate vial se da prioridad al menor costo de apertura.
            return (sum(min(self.ins.d(d,c),self.ins.d(c,d))
                        for c in self.ins.clients),
                    self.ins.depots[d].opening_cost, d)
        closed.sort(key=target_rank)
        if not small:
            closed=closed[:self.open_cfg.max_depots_large]
        for target in closed:
            groups,truncated=self._candidate_groups(target)
            if truncated:
                self.opening_incomplete=True
            for group in groups:
                if (self.open_cfg.max_evaluations and
                    self.tried['opening']>=self.open_cfg.max_evaluations):
                    self.stage_exhausted.add('opening')
                    self.opening_limit_reason='LIMITE_EVALUACIONES_APERTURA'
                    return False
                if (self.open_cfg.max_seconds and
                    _perf_counter_ii()-self.opening_started>=self.open_cfg.max_seconds):
                    self.stage_exhausted.add('opening')
                    self.opening_limit_reason='LIMITE_TIEMPO_APERTURA'
                    return False
                key=(tuple(sorted(self.current.open_depots)),target,group,
                     tuple(sorted(self.current.assignment.items())))
                if key in self._opening_candidates_seen:
                    continue
                self._opening_candidates_seen.add(key)
                demand=sum(self.ins.clients[c].delivery for c,x in self.current.assignment.items() if x==target)
                demand+=sum(self.ins.clients[c].delivery for c in group)
                if demand>self.ins.depots[target].capacity+EPS:
                    continue
                self.tried['opening']+=1
                trial,affected=self._reconstruct_group(target,group)
                if trial is None or trial.total_cost>=self.current.total_cost-self.cfg.improvement_eps:
                    continue
                if (self._mejor_barrido is not None and
                    trial.total_cost>=self._mejor_barrido.total_cost-self.cfg.improvement_eps):
                    continue
                errors=validate_solution(self.ins,trial)
                if errors:
                    raise AssertionError('Apertura inválida: '+'; '.join(errors[:12]))
                self._mejor_barrido=trial
                self._detalle_mejor=('opening',f'abrir {target}; mover {" | ".join(group)}',tuple(sorted(affected)))
        return False

    def optimize(self):
        # El orden y la regla de mejor mejora POR VECINDARIO son los originales.
        stages=[('2-opt','two_opt',self.two_opt),
                ('Relocate','relocate',self.relocate),
                ('Swap','swap',self.swap),
                ('Fusión','fusion',self.fusion),
                ('Reasignación','reassignment',self.reassignment)]
        if self.open_cfg.enabled:
            stages.append(('Apertura','opening',self.apertura_localizacion))
        for k in range(1,self.cfg.max_passes+1):
            self.pass_number=k
            any_change=False
            for label,key,search in stages:
                # La apertura tiene presupuesto propio; no consume el de los
                # otros vecindarios, que siguen las cuotas originales.
                def can_continue():
                    return key not in self.stage_exhausted and (key=='opening' or self._in_budget())
                while can_continue():
                    self._mejor_barrido=None
                    self._detalle_mejor=None
                    ev_before=self.tried.get(key,0)
                    ac_before=self.accepted_by_stage.get(key,0)
                    t0=_perf_counter_ii()
                    search()
                    improved=self._aceptar_mejor_barrido()
                    elapsed=_perf_counter_ii()-t0
                    self.times_by_stage.append(dict(pasada=k,movimiento=label,
                        evaluaciones=self.tried[key]-ev_before,
                        mejoras=self.accepted_by_stage[key]-ac_before,
                        segundos=elapsed,
                        completo=(key not in self.stage_exhausted and not
                                  (key=='opening' and self.opening_incomplete))))
                    if not improved:break
                    any_change=True
                    improvement=self.history[-1]
                    print('    {}: ahorro {:,.2f} COP; costo {:,.2f} COP'.format(
                        label,improvement.saving,improvement.after_cost),flush=True)
                if self.exhausted and key!='opening':
                    # Aún se revisa Apertura, cuya cuota es independiente.
                    continue
            if not any_change or self.exhausted:
                # El criterio del optimizador original para los movimientos
                # previos se conserva; no se declara convergencia si hubo topes.
                break
        result=copy.deepcopy(self.current)
        result.feasible=True
        result.violations=validate_solution(self.ins,result)
        if result.violations:
            raise AssertionError('Búsqueda local inválida: '+'; '.join(result.violations[:12]))
        if result.total_cost>self.initial.total_cost+self.cfg.improvement_eps:
            raise AssertionError('Fase II aumentó el costo diario')
        return result

def _solo_oficial(sol: Solution, label: str, config: SearchConfig,
                  extra: dict | None = None) -> RunResult:
    """Vista de UN resultado para el Excel, no borra el top cinco en RAM."""
    params = {f.name: getattr(config, f.name) for f in fields(config)}
    params.update(extra or {})
    params['fase'] = label
    single = copy.deepcopy(sol)
    return RunResult(0, 0, 0, [], [], {single.config_id:copy.deepcopy(single)},
                     single, [copy.deepcopy(single)], params, [], 0.0,
                     'Resultado oficial de '+label)


def _excel_memoria(ins: Instance, result: RunResult) -> BytesIO:
    """Genera el esquema Excel original directamente en RAM, sin archivo temporal."""
    data = BytesIO()
    export_results_excel(data, ins, result)
    data.seek(0)
    return data


def _datos_excel(book: BytesIO) -> dict[str, pd.DataFrame]:
    book.seek(0)
    return pd.read_excel(book,sheet_name=None,engine='openpyxl',dtype=object)


def recuperar_excel_consolidado(path: str, phase: str, ins: Instance):
    """Recuperar desde libro consolidado la Fase I o II, con su validador."""
    if phase not in ('FI','FII'):
        raise ValueError('phase debe ser FI o FII')
    data = pd.read_excel(path,sheet_name=None,engine='openpyxl',dtype=object)
    prefix=phase+'_'
    group={name[len(prefix):]:df for name,df in data.items()
           if name.startswith(prefix)}
    missing=set(COLUMNS)-set(group)
    # Se omiten del libro consolidado las tablas vacías de intentos/evolución.
    # Se crean solo temporalmente en RAM para reutilizar el importador original.
    optional={'Arranques','Combinaciones','Evolucion'}
    if missing-optional:
        raise ValueError(f'Faltan hojas de {phase}: {sorted(missing-optional)}')
    buff=BytesIO()
    with pd.ExcelWriter(buff,engine='openpyxl') as writer:
        for name, df in group.items():
            df.to_excel(writer,sheet_name=name,index=False)
        for name in sorted(missing & optional):
            pd.DataFrame(columns=COLUMNS[name]).to_excel(writer,sheet_name=name,index=False)
    buff.seek(0)
    return read_results_excel(buff,ins)


def exportar_excel_integrado(path: str, ins: Instance,
                              first: Solution, final: Solution,
                              base: SearchConfig, metrics: dict,
                              candidates: list[dict], moves: list[dict],
                              comparison: dict, *, top5=None,
                              lotes=None, tiempos=None, ciclos=None,
                              evolucion_fase1=None, ahorro_etapas=None) -> str:
    """Excel único con UNA solución oficial por fase y trazabilidad verificada."""
    first_sheets=_datos_excel(_excel_memoria(ins,_solo_oficial(first,'I',base,metrics)))
    final_sheets=_datos_excel(_excel_memoria(ins,_solo_oficial(final,'II',base,metrics)))
    path=str(Path(path).resolve())
    Path(path).parent.mkdir(parents=True,exist_ok=True)
    with pd.ExcelWriter(path,engine='openpyxl') as writer:
        for prefix,book in [('FI_',first_sheets),('FII_',final_sheets)]:
            for name,data in book.items():
                if not data.empty:
                    data.to_excel(writer,sheet_name=prefix+name,index=False)
        pd.DataFrame([comparison]).to_excel(writer,sheet_name='Comparacion_Fases',index=False)
        pd.DataFrame([metrics]).to_excel(writer,sheet_name='Control_Ejecucion',index=False)
        extras=[('Top5_FaseI',top5),('Evaluacion_Candidatos',candidates),
                ('FaseI_Lotes',lotes),('Movimientos_FaseII',moves),
                ('Tiempos_Movimientos',tiempos),('Ciclos_Adicionales',ciclos),
                ('FaseI_Evolucion',evolucion_fase1),('Ahorros_Etapas',ahorro_etapas)]
        for name,rows in extras:
            if rows:
                pd.DataFrame(rows).to_excel(writer,sheet_name=name,index=False)
        coords=[dict(nodo_id=d.id,lat=d.y,lon=d.x,tipo='BODEGA')
                for d in ins.depots.values() if d.x is not None and d.y is not None]
        coords.extend(dict(nodo_id=c.id,lat=c.y,lon=c.x,tipo='CLIENTE')
                for c in ins.clients.values() if c.x is not None and c.y is not None)
        if coords:
            pd.DataFrame(coords).to_excel(writer,sheet_name='Coordenadas_Nodos',index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes='A2'
            ws.auto_filter.ref=ws.dimensions
            ws.sheet_view.showGridLines=False
            for cell in ws[1]:
                cell.fill=PatternFill('solid',fgColor='17365D')
                cell.font=Font(color='FFFFFF',bold=True,size=10)
                cell.alignment=Alignment(vertical='center',wrap_text=True)
            for col in ws.columns:
                first_cell=col[0]
                n=max((len(str(c.value or '')) for c in list(col)[:100]),default=12)
                ws.column_dimensions[first_cell.column_letter].width=min(45,max(13,n+2))
    for phase,expected in [('FI',first),('FII',final)]:
        rec=recuperar_excel_consolidado(path,phase,ins)
        restored=rec['best_global']
        if restored is None or len(rec['solutions'])!=1 or (
            signature(restored)!=signature(expected) or
            abs(restored.total_cost-expected.total_cost)>1e-5 or
            validate_solution(ins,restored)):
            raise AssertionError('Reimportación de '+phase+' no coincide con solución validada')
    return path


def imprimir_oficial(titulo: str, sol: Solution):
    print('\n'+('='*18)+' '+titulo+' '+('='*18))
    print('COSTOS DIARIOS: bodegas {:,.2f} | arriendo de {} vehículos {:,.2f} | '
          'transporte {:,.2f} | TOTAL {:,.2f} COP/día'.format(
        sol.opening_cost,len(sol.routes),sol.vehicle_rental_cost,sol.transport_cost,sol.total_cost))
    print('Bodegas:',', '.join(sol.open_depots), '| Rutas:',len(sol.routes),
          '| Distancia: {:.2f} km'.format(sol.distance_km))
    print('ASIGNACIONES (cliente -> bodega):')
    for c,d in sorted(sol.assignment.items()):
        print(f'  {c}: {d}')
    print('RUTAS COMPLETAS (1 ruta = 1 vehículo diario que sale y retorna):')
    alquiler_por_ruta=(sol.vehicle_rental_cost / len(sol.routes) if sol.routes else 0.0)
    for r in sorted(sol.routes,key=lambda z:(z.depot_id,z.id)):
        print('  {} [{}] {:.2f} km, {:.2f} min, transporte {:,.2f}, '
              'arriendo {:,.2f}, total ruta {:,.2f} COP/día, carga inicial {:.0f}'.format(
            r.id,r.depot_id,r.km,r.minutes,r.cost,alquiler_por_ruta,
            r.cost+alquiler_por_ruta,r.initial_load))
        for st in r.stops:
            print('    {:02d} {:>4} {:<12} unidades {:>5.0f} carga {:>5.0f}->{:>5.0f} '
                  'arco {:.2f} km acum {:.2f} km'.format(
                st.order,st.node_id,st.operation,st.amount,st.load_before,
                st.load_after,st.leg_km,st.cumulative_km))


# ========== FASE I: LOTES POR PAREJA, ESTANCAMIENTO REPRODUCIBLE ==========
def construir_hasta_estancamiento(ins: Instance, base: SearchConfig,
        initial_seeds: tuple[int, ...], ratios: tuple[float, ...],
        collector: CincoMejoresFaseI, stagnation_batches: int = 2,
        max_batches: int | None = None, max_total_constructions: int | None = None,
        constructions_per_pair: int = 72):
    """Ejecuta N construcciones POR combinación (semilla, cierre) en cada lote.

    El lote completo consta de len(semillas) * len(cierres) * N intentos.
    Los siguientes lotes usan semillas distintas, y el contador de lotes
    estancados vuelve a cero al mejorar el mejor costo GLOBAL validado.
    Un límite de seguridad jamás se interpreta como convergencia matemática.
    """
    if not initial_seeds or len(set(initial_seeds)) != len(initial_seeds):
        raise ValueError('Las semillas iniciales deben ser distintas')
    if not ratios or len(set(ratios)) != len(ratios):
        raise ValueError('Los cierres deben ser distintos')
    if (min(stagnation_batches, constructions_per_pair) < 1 or
        (max_batches is not None and max_batches < 1) or
        (max_total_constructions is not None and max_total_constructions < 1)):
        raise ValueError('Los límites de lotes deben ser positivos')
    seeds_per_batch=len(initial_seeds)
    pair_count=seeds_per_batch*len(ratios)
    batch_size=constructions_per_pair * pair_count
    # La cantidad por pareja no se reparte entre los cruces; cada uno recibe
    # exactamente la misma cuota, incluidas las semillas de los lotes nuevos.
    start_time=perf_counter()
    seen_seed=set()
    seen_deterministic=set()
    logs=[]
    # Instrumentación pasiva: no participa en construcción, ranking ni estancamiento.
    evolucion_fase1=[]
    primera_factible_costo=None
    mejor_observado=math.inf
    observaciones_factibles=0
    total_attempts=0
    total_feasible=0
    best_cost=math.inf
    failed_batches=0
    stop_reason='LIMITE_SEGURIDAD_LOTES'
    next_seed=max(initial_seeds)+1
    batch = 0
    while failed_batches < stagnation_batches:
        batch += 1
        if max_batches is not None and batch > max_batches:
            stop_reason = 'LIMITE_SEGURIDAD_LOTES'
            break
        if (max_total_constructions is not None and
                total_attempts + batch_size > max_total_constructions):
            stop_reason='LIMITE_SEGURIDAD_CONSTRUCCIONES_SIN_LOTE_COMPLETO'
            break
        if batch == 1:
            seeds=initial_seeds
        else:
            seeds=tuple(range(next_seed,next_seed+seeds_per_batch))
            next_seed+=seeds_per_batch
        if any(seed in seen_seed for seed in seeds):
            raise AssertionError('Semilla principal reutilizada en los lotes')
        seen_seed.update(seeds)
        beginning=perf_counter()
        before=best_cost
        batch_feasible=0
        batch_attempts=0
        pairs=[(seed,ratio) for seed in seeds for ratio in ratios]
        for seed,ratio in pairs:
            budget=constructions_per_pair
            cfg=replace(base,execution_mode='budget',construction_budget=budget,
                        seed=seed,geometric_p=0.95,cluster_split_ratio=ratio)
            # Recoger solo mejoras estrictas del incumbente observado. El orden
            # de llamada al recolector original y las semillas no cambian.
            def registrar_factible(sol):
                nonlocal primera_factible_costo, mejor_observado, observaciones_factibles
                observaciones_factibles += 1
                if primera_factible_costo is None:
                    primera_factible_costo=float(sol.total_cost)
                if sol.total_cost < mejor_observado-1e-6:
                    previo=mejor_observado
                    mejor_observado=float(sol.total_cost)
                    evolucion_fase1.append(dict(
                        orden=len(evolucion_fase1)+1, lote=batch,
                        semilla_principal=seed, cierre=ratio,
                        factibles_observadas=observaciones_factibles,
                        costo_anterior=(previo if math.isfinite(previo) else None),
                        mejor_costo=mejor_observado,
                        ahorro=(previo-mejor_observado if math.isfinite(previo) else 0.0),
                        segundos_desde_inicio=perf_counter()-start_time))
                collector.incluir(sol)
            result=multi_start(ins,cfg,on_feasible=registrar_factible,
                               deterministic_seen=seen_deterministic,
                               explore_extended=(batch > 1))
            batch_attempts+=len(result.attempts)
            batch_feasible+=sum(a.feasible for a in result.attempts)
            # La mejor se conserva por copia; no se exportan rutas secundarias.
        total_attempts+=batch_attempts
        total_feasible+=batch_feasible
        elites=collector.finales()
        if elites:
            best_cost=min(x.total_cost for x in elites)
        improved=best_cost < before - 1e-6
        failed_batches=0 if improved else failed_batches+1
        logs.append(dict(lote=batch,semillas=','.join(map(str,seeds)),
              construcciones=batch_attempts, factibles=batch_feasible,
              mejor_costo=best_cost if math.isfinite(best_cost) else None,
              mejoro=improved, consecutivos_sin_mejora=failed_batches,
              segundos=perf_counter()-beginning))
        print('FASE I lote {}: {:d} construcciones, {:d} factibles, '
              'mejor {:,.0f} COP, sin mejora {}/{}'.format(batch,batch_attempts,
              batch_feasible,best_cost,failed_batches,stagnation_batches),flush=True)
        if batch_attempts < batch_size:
            stop_reason='LOTE_INCOMPLETO_SIN_CONFIGURACIONES_EVALUABLES'
            break
        if failed_batches >= stagnation_batches:
            stop_reason='ESTANCAMIENTO_{}_LOTES_COMPLETOS'.format(stagnation_batches)
            break
    elites=collector.finales()
    if not elites:
        raise RuntimeError('No se encontró ninguna solución factible validada en la Fase I')
    if len(elites) < collector.limit:
        print(f'AVISO: solo se encontraron {len(elites)} soluciones distintas; se refinan todas.', flush=True)
    # La colección contiene copias profundas validadas.
    return elites, dict(lotes=len(logs),resumen_lotes=logs,
           construcciones_totales=total_attempts, factibles_totales=total_feasible,
           estructuras_distintas=collector.distintas,
           primera_factible_costo=primera_factible_costo,
           evolucion_fase1=evolucion_fase1,
           estancamiento_lotes=stagnation_batches,
           lotes_sin_mejora=failed_batches,
           motivo_terminacion=stop_reason,
           segundos=perf_counter()-start_time,
           semillas_iniciales=','.join(map(str,initial_seeds)),
           cierres=','.join(map(str,ratios)),
           construcciones_por_pareja=constructions_per_pair,
           construcciones_por_lote=batch_size,
           parejas_por_lote=pair_count,
           configuraciones_deterministas_unicas=len(seen_deterministic))


def _status_local(opt: BusquedaLocalMejorMejora, cfg: LocalConfig) -> str:
    if getattr(opt,'opening_limit_reason',''):
        return 'BUSQUEDA_INCOMPLETA_'+opt.opening_limit_reason
    if getattr(opt,'opening_incomplete',False):
        return 'EXPLORACION_CONTROLADA_APERTURA_NO_EXHAUSTIVA'
    if opt.exhausted or opt.count >= cfg.max_evaluations:
        return 'BUSQUEDA_INCOMPLETA_LIMITE_GLOBAL'
    if opt.stage_exhausted:
        return 'BUSQUEDA_INCOMPLETA_LIMITE_MOVIMIENTO_'+','.join(sorted(opt.stage_exhausted))
    if opt.pass_number >= cfg.max_passes and len(opt.history):
        return 'LIMITE_PASADAS_NO_CONVERGENCIA_GARANTIZADA'
    return 'SIN_MEJORA_EN_BARRIDO_EXPLORADO'

def ejecutar_integrado(argv=None):
    """Fase I por lotes; Fase II original de tres ciclos y mejor mejora POR vecindario.

    Sin Or-opt ni selección global; movimiento acotado de apertura adicional.
    Si un límite de seguridad se agota, la ejecución se clasifica como incompleta.
    """
    ap=_parser_integrado.ArgumentParser(description='LogiData: construcción por lotes + búsqueda local original de tres ciclos')
    for name in ('bodegas','clientes','distancias','tiempos'):
        ap.add_argument('--'+name,default=FILES_DEFAULT[name])
    ap.add_argument('--seeds',default=None,help='Semillas del primer lote, separadas por comas')
    ap.add_argument('--seed',type=int,default=4110)
    ap.add_argument('--num-seeds',type=int,default=DEFAULT_MAIN_SEED_COUNT)
    ap.add_argument('--splits',default=','.join(map(str,MAIN_SPLIT_LEVELS)))
    ap.add_argument('--top',type=int,default=5)
    ap.add_argument('--depot-cost-basis',choices=('investment','daily'),default='investment',
                    help='investment: amortizar costo de entrada 15*365; daily: entrada ya expresada por día')
    ap.add_argument('--daily-vehicle-rent',type=float,default=147572.55,
                    help='Arriendo COP/día por Renault Trafic L2H2 (una vez por ruta)')
    ap.add_argument('--constructions-per-pair',type=int,default=72,
                    help='72 construcciones por CADA par (semilla,cierre), en cada lote')
    ap.add_argument('--stagnation-batches',type=int,default=2)
    ap.add_argument('--max-batches',type=int,default=None,
                    help='Protección opcional; su agotamiento no significa estancamiento')
    ap.add_argument('--max-total-constructions',type=int,default=None,
                    help='Protección opcional: solo permite lotes completos')
    ap.add_argument('--additional-cycles',type=int,default=3,
                    help='Hasta tres ciclos adicionales desde la mejor Fase II, presupuestos renovados')
    ap.add_argument('--max-passes',type=int,default=4)
    ap.add_argument('--max-evaluations',type=int,default=30000)
    ap.add_argument('--max-accepted',type=int,default=150)
    ap.add_argument('--max-stage-evaluations',type=int,default=7000)
    ap.add_argument('--max-stage-accepted',type=int,default=30)
    ap.add_argument('--candidate-insertions',type=int,default=4,
                    help='Las cuatro posiciones originales por delta vial dirigido')
    ap.add_argument('--disable-opening',action='store_true',help='Desactivar solo el nuevo movimiento')
    ap.add_argument('--opening-small-threshold',type=int,default=5)
    ap.add_argument('--opening-max-groups',type=int,default=28)
    ap.add_argument('--opening-max-depots',type=int,default=6)
    ap.add_argument('--opening-max-evaluations',type=int,default=None,
                    help='0 desactiva el tope; para 5 clientes permite barrido completo')
    ap.add_argument('--opening-max-accepted',type=int,default=20)
    ap.add_argument('--opening-max-seconds',type=float,default=0.0)
    ap.add_argument('--output',default='LogiData_Fases_I_II_consolidado.xlsx')
    args=ap.parse_args(argv)
    if args.top!=5:
        ap.error('Se conserva la selección de cinco candidatas (top=5)')
    if args.additional_cycles<0:
        ap.error('--additional-cycles debe ser >= 0')
    seeds=(tuple(int(x.strip()) for x in args.seeds.split(',')) if args.seeds
           else tuple(args.seed+i for i in range(args.num_seeds)))
    ratios=tuple(float(x.strip()) for x in args.splits.split(','))
    ins,base,input_cfg=setup_instance({k:getattr(args,k) for k in
                               ('bodegas','clientes','distancias','tiempos')},
                               depot_cost_basis=args.depot_cost_basis,
                               vehicle_rental_daily=args.daily_vehicle_rent)
    base=replace(base,construction_budget=args.constructions_per_pair,
                 geometric_p=0.95,keep_top=10)
    collector=CincoMejoresFaseI(ins,5)
    began=perf_counter()
    print('FASE I: {} construcciones por pareja y lote, {} parejas, estancamiento {} lotes'.format(
        args.constructions_per_pair,len(seeds)*len(ratios),args.stagnation_batches),flush=True)
    elites,fi=construir_hasta_estancamiento(ins,base,seeds,ratios,collector,
        stagnation_batches=args.stagnation_batches,max_batches=args.max_batches,
        max_total_constructions=args.max_total_constructions,
        constructions_per_pair=args.constructions_per_pair)
    initial=copy.deepcopy(elites[0]);initial.id='FI-OFICIAL'
    if validate_solution(ins,initial):
        raise AssertionError('La mejor solución de Fase I no es factible')
    lc=LocalConfig(args.max_passes,args.max_evaluations,args.max_accepted,
        args.candidate_insertions,max_stage_evaluations=args.max_stage_evaluations,
        max_stage_accepted=args.max_stage_accepted)
    opening_budget=(0 if len(ins.clients)<=args.opening_small_threshold else 350) if args.opening_max_evaluations is None else args.opening_max_evaluations
    opening_cfg=OpeningConfig(not args.disable_opening,args.opening_small_threshold,
        args.opening_max_groups,args.opening_max_depots,opening_budget,
        args.opening_max_accepted,args.opening_max_seconds)
    opening_cfg.validate()
    evaluated=[]; all_moves=[]; all_timing=[]; global_move_serial=0
    def save_instrumentation(opt,label):
        nonlocal global_move_serial
        for event in opt.economic_moves:
            global_move_serial+=1
            all_moves.append(dict(candidato=label,orden_global=global_move_serial,
                **event))
        for event in opt.times_by_stage:
            all_timing.append(dict(candidato=label,**event))

    for rank,start in enumerate(elites,1):
        print('FASE II: candidata {}/{}; inicial {:,.0f} COP'.format(
            rank,len(elites),start.total_cost),flush=True)
        optimizer=BusquedaLocalConApertura(ins,start,lc,opening_cfg,base)
        result=optimizer.optimize()
        optimizer.elapsed_seconds=perf_counter()-optimizer.started
        result.id=f'FII-CAND-{rank:02d}'
        violations=validate_solution(ins,result)
        if violations: raise AssertionError('Fase II inválida: '+'; '.join(violations[:10]))
        if not set(start.open_depots).issubset(set(result.open_depots)):
            raise AssertionError('Se cerró una bodega durante Fase II')
        if result.total_cost>start.total_cost+lc.improvement_eps:
            raise AssertionError('Fase II empeoró la candidata inicial')
        evaluated.append((start,result,optimizer))
        save_instrumentation(optimizer,str(rank))
        print('  Refinada {:,.0f}; {} mejoras, {} evaluaciones'.format(
             result.total_cost,len(optimizer.history),optimizer.count),flush=True)
    winner=min(evaluated,key=lambda v:(v[1].total_cost,v[1].distance_km,
                                         len(v[1].routes),signature(v[1])))
    # Los ciclos adicionales recuperan EXACTAMENTE el comportamiento anterior:
    # una sola pasada por ciclo, presupuestos completos y reinicio de clases/contadores.
    best=copy.deepcopy(winner[1]); extras=[]
    extra_cfg=replace(lc,max_passes=1)
    for cycle in range(1,args.additional_cycles+1):
        start=copy.deepcopy(best)
        opt=BusquedaLocalConApertura(ins,start,extra_cfg,opening_cfg,base)
        changed=opt.optimize()
        opt.elapsed_seconds=perf_counter()-opt.started
        failures=validate_solution(ins,changed)
        if failures: raise AssertionError('Ciclo adicional inválido: '+'; '.join(failures[:10]))
        if not set(start.open_depots).issubset(set(changed.open_depots)):
            raise AssertionError('Se cerró una bodega durante un ciclo adicional')
        if changed.total_cost>start.total_cost+lc.improvement_eps:
            raise AssertionError('El ciclo adicional empeoró su solución inicial')
        saving=start.total_cost-changed.total_cost
        extra_status=_status_local(opt,extra_cfg)
        extras.append(dict(ciclo=cycle,inicial=start.total_cost,final=changed.total_cost,
            ahorro=saving,evaluaciones=opt.count+opt.tried['opening'],mejoras=len(opt.history),
            evaluaciones_apertura=opt.tried['opening'],aperturas_aceptadas=opt.accepted_by_stage['opening'],
            motivo=extra_status,segundos=opt.elapsed_seconds))
        save_instrumentation(opt,'ADICIONAL-{:02d}'.format(cycle))
        if saving <= lc.improvement_eps:
            print('Ciclo adicional {}: sin ahorro; fin, motivo {}'.format(cycle,extra_status),flush=True)
            break
        best=copy.deepcopy(changed)
        print('Ciclo adicional {}: ahorro {:,.0f} COP'.format(cycle,saving),flush=True)
    final=copy.deepcopy(min((initial,winner[1],best),key=lambda s:
         (s.total_cost,s.distance_km,len(s.routes),signature(s))))
    final.id='FII-OFICIAL'
    if validate_solution(ins,final):
        raise AssertionError('La solución final no superó la validación')
    if final.total_cost>initial.total_cost+lc.improvement_eps:
        raise AssertionError('El costo total de Fase II aumentó')
    stats=[]
    for rank,(before,after,opt) in enumerate(evaluated,1):
        stats.append(dict(candidato=rank,inicial=before.total_cost,final=after.total_cost,
            ahorro=before.total_cost-after.total_cost,rutas_inicial=len(before.routes),
            rutas_final=len(after.routes),bodegas=' | '.join(after.open_depots),
            mejoras=len(opt.history),evaluaciones=opt.count+opt.tried['opening'],
            evaluaciones_apertura=opt.tried['opening'],
            aperturas_aceptadas=opt.accepted_by_stage['opening'],
            motivo_terminacion=_status_local(opt,lc),segundos=opt.elapsed_seconds,
            aperturas=' | '.join(sorted(set(after.open_depots)-set(before.open_depots))),
            kilometros_iniciales=before.distance_km,kilometros_finales=after.distance_km,
            motivo_apertura=opt.opening_limit_reason or ('EXPLORACION_CONTROLADA' if opt.opening_incomplete else 'COMPLETA'),
            aceptadas_por_movimiento=' | '.join(f'{k}:{v}' for k,v in sorted(opt.accepted_by_stage.items()))))
    summary=dict(costo_inicial=initial.total_cost,costo_final=final.total_cost,
        bodegas_diario_fase1=initial.opening_cost,bodegas_diario_fase2=final.opening_cost,
        arriendo_diario_fase1=initial.vehicle_rental_cost,arriendo_diario_fase2=final.vehicle_rental_cost,
        transporte_diario_fase1=initial.transport_cost,transporte_diario_fase2=final.transport_cost,
        vehiculos_fase1=len(initial.routes),vehiculos_fase2=len(final.routes),
        ahorro_cop=initial.total_cost-final.total_cost,
        ahorro_porcentaje=100*(initial.total_cost-final.total_cost)/initial.total_cost,
        rutas_fase1=len(initial.routes),rutas_fase2=len(final.routes),
        km_fase1=initial.distance_km,km_fase2=final.distance_km,
        bodegas_fase1=' | '.join(initial.open_depots),
        bodegas_fase2=' | '.join(final.open_depots),
        mejor_candidato_fase2=next(i for i,(_,s,_) in enumerate(evaluated,1) if s is winner[1]),
        ciclos_adicionales_ejecutados=len(extras),
        ahorros_ciclos_adicionales=sum(c['ahorro'] for c in extras))
    metrics=dict(base_costos_bodegas=args.depot_cost_basis,
        arriendo_diario_vehiculo=ins.vehicle_rental_daily,
        costo_diario_bodegas_final=final.opening_cost,
        costo_diario_arriendo_final=final.vehicle_rental_cost,
        costo_diario_transporte_final=final.transport_cost,
        vehiculos_diarios_final=len(final.routes),
        semillas_iniciales=','.join(map(str,seeds)),
        cierres=','.join(map(str,ratios)),
        geometric_p=0.95,construcciones_por_pareja=args.constructions_per_pair,
        construcciones_por_lote=fi['construcciones_por_lote'],
        lotes_fase1=fi['lotes'],construcciones_fase1=fi['construcciones_totales'],
        construcciones_factibles_fase1=fi['factibles_totales'],
        estructuras_distintas_fase1=collector.distintas,
        motivo_terminacion_fase1=fi['motivo_terminacion'],
        lotes_sin_mejora=fi['lotes_sin_mejora'],
        estrategia_fase2='2-opt > Relocate > Swap > Fusion > Reasignacion > Apertura (opcional), mejor mejora por vecindario',
        apertura_habilitada=opening_cfg.enabled,
        limite_evaluaciones_apertura=opening_cfg.max_evaluations,
        or_opt_2='NO (recuperada version 3 ciclos)',
        ciclos_adicionales_limite=args.additional_cycles,
        ciclos_adicionales_realizados=len(extras),
        motivo_ciclos_adicionales=(extras[-1]['motivo'] if extras else 'DESACTIVADOS'),
        evaluaciones_fase2=sum(s['evaluaciones'] for s in stats)+sum(c['evaluaciones'] for c in extras),
        evaluaciones_apertura=sum(s['evaluaciones_apertura'] for s in stats)+sum(c['evaluaciones_apertura'] for c in extras),
        mejoras_aceptadas_fase2=sum(s['mejoras'] for s in stats)+sum(c['mejoras'] for c in extras),
        motivo_terminacion_fase2=' | '.join(f'{s["candidato"]}:{s["motivo_terminacion"]}' for s in stats),
        tiempo_fase1_segundos=fi['segundos'],
        tiempo_fase2_segundos=perf_counter()-began-fi['segundos'],
        costo_cop_km=ins.cost_per_km,
        mejoras_fase2=sum(s['mejoras'] for s in stats)+sum(c['mejoras'] for c in extras),
        tiempo_total_segundos=perf_counter()-began)
    imprimir_oficial('FASE I OFICIAL',initial)
    imprimir_oficial('FASE II OFICIAL',final)
    print('COMPARACION FI/FII: {:,.0f} -> {:,.0f} COP; ahorro {:,.0f}; rutas {} -> {}'.format(
        initial.total_cost,final.total_cost,summary['ahorro_cop'],
        len(initial.routes),len(final.routes)),flush=True)
    print('MOTIVOS FASE I:',fi['motivo_terminacion'],flush=True)
    print('MOTIVOS FASE II:',metrics['motivo_terminacion_fase2'],flush=True)
    summary['minutos_fase1']=sum(r.minutes for r in initial.routes)
    summary['minutos_fase2']=sum(r.minutes for r in final.routes)
    # Marcar resultados limitados como incompletos, nunca óptimo certificado.
    def stage_status(opt,cfg):
        if opt.opening_limit_reason:
            return 'BUSQUEDA_INCOMPLETA_'+opt.opening_limit_reason
        if opt.opening_incomplete:
            return 'EXPLORACION_CONTROLADA_APERTURA_SIN_OPTIMALIDAD_CERTIFICADA'
        return _status_local(opt,cfg)
    metrics['motivo_terminacion_fase2']=' | '.join(
        f'{i}:{stage_status(opt,lc)}' for i,(_,_,opt) in enumerate(evaluated,1))
    # El ahorro total se mide desde la primera solución factible de Fase I.
    # No se suma el ahorro del candidato ganador a la mejora global de Fase I:
    # ese candidato puede ser distinto de la mejor solución constructiva.
    first_feasible=fi.get('primera_factible_costo')
    saved_build=(first_feasible-initial.total_cost if first_feasible is not None else None)
    saved_local=initial.total_cost-final.total_cost
    saved_total=(first_feasible-final.total_cost if first_feasible is not None else None)
    ahorro_etapas=[
        dict(etapa='Fase I · construcciones',costo_inicial=first_feasible,
             costo_final=initial.total_cost,ahorro=saved_build,
             porcentaje=(100*saved_build/first_feasible if first_feasible else None)),
        dict(etapa='Fase II · búsqueda local',costo_inicial=initial.total_cost,
             costo_final=final.total_cost,ahorro=saved_local,
             porcentaje=(100*saved_local/initial.total_cost if initial.total_cost else None)),
        dict(etapa='Total de la ejecución',costo_inicial=first_feasible,
             costo_final=final.total_cost,ahorro=saved_total,
             porcentaje=(100*saved_total/first_feasible if first_feasible else None))]
    metrics['primera_factible_costo_fase1']=first_feasible
    metrics['ahorro_construccion_fase1']=saved_build
    metrics['ahorro_busqueda_local_fase2']=saved_local
    metrics['ahorro_total_desde_primera_factible']=saved_total
    # Normalizar porcentajes y ahorro acumulado por trayectoria de cada
    # candidato o ciclo; un candidato nunca se acumula con otro.
    acum_por_candidato={}
    for event in all_moves:
        label=str(event['candidato'])
        acum_por_candidato[label]=acum_por_candidato.get(label,0.0)+float(event['ahorro'])
        event['ahorro_acumulado_candidato']=acum_por_candidato[label]
    top_data=[dict(posicion=i,costo_total=s.total_cost,bodegas=' | '.join(s.open_depots),
        rutas=len(s.routes),vehiculos=len(s.routes),km=s.distance_km,
        costo_bodegas=s.opening_cost,arriendo=s.vehicle_rental_cost,
        transporte=s.transport_cost,semilla=s.seed) for i,s in enumerate(elites,1)]
    path=exportar_excel_integrado(args.output,ins,initial,final,base,metrics,
        stats,all_moves,summary,top5=top_data,lotes=fi['resumen_lotes'],
        tiempos=all_timing,ciclos=extras,
        evolucion_fase1=fi['evolucion_fase1'],ahorro_etapas=ahorro_etapas)
    print('Excel consolidado:',path,flush=True)
    return initial,final,path


# ============================================================================
# DASHBOARD STREAMLIT Y PLOTLY (INTEGRADO EN ESTE MISMO ARCHIVO)
# Solo se ejecuta después de terminar la heurística o con --view.
# ============================================================================
def _logidata_dashboard():
    #!/usr/bin/env python3
    # -*- coding: utf-8 -*-
    """Dashboard LogiData: visualizador independiente (no ejecuta la heurística).

    Inicio: streamlit run LogiData_Dashboard.py
    Lee tanto el Excel histórico (solo FII detallada) como el Excel instrumentado
    (con FI y FII detalladas y trazabilidad mínima). No inventa valores ausentes.
    """

    from collections import Counter
    from io import BytesIO
    from pathlib import Path
    import math
    import re
    import sys

    import numpy as np
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
    import streamlit as st

    st.set_page_config(page_title='LogiData | Centro de resultados', page_icon='🔷',
                       layout='wide', initial_sidebar_state='expanded')

    NAVY = '#17365E'
    TEAL = '#138F8A'
    BLUE = '#2865B8'
    ORANGE = '#E29B42'
    VIOLET = '#7566C8'
    RED = '#CC5261'
    PALETTE = ['#138F8A', '#2865B8', '#E29B42', '#7566C8', '#CE557B',
               '#3987BF', '#6C9E54', '#B78038', '#4E669F', '#A066B0', '#6B907D', '#3B6C73']

    st.markdown('''<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    :root {--ld-ink:#17365E;--ld-soft:#60738F;--ld-blue:#2865B8;--ld-border:#DEE8F5}
    html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
       font-family:Inter,system-ui,-apple-system,Segoe UI,Arial,sans-serif}
    [data-testid="stAppViewContainer"] {background:#F6F9FD;color:#19304C}
    .block-container {padding-top:1.0rem;padding-bottom:2.1rem;max-width:1520px}
    html,body {font-size:16px}
    [data-testid="stMarkdownContainer"] {font-size:1.04rem}
    [data-testid="stWidgetLabel"] p {font-size:1.04rem !important}
    [data-testid="stSelectbox"] * {font-size:1.04rem}
    h1,h2,h3,h4 {color:#172642;letter-spacing:-.025em}
    h1 {font-weight:800} h2 {font-weight:750} h3 {font-weight:700}
    [data-testid="stSidebar"] {background:linear-gradient(170deg,#10284B,#183A66 72%,#1B477B);
       border-right:1px solid #29547E}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p {color:#EDF0FC !important}
    [data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"]>div {
        background:#FFFFFF;border:1px solid #D9DDEA;border-radius:12px}
    [data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] * {
        color:#172642}
    [data-testid="stSidebar"] hr {border-color:#416385}
    /* Botones azules, incluidos los de descarga, con contraste y estados visibles. */
    [data-testid="stButton"] button,
    [data-testid="stDownloadButton"] button,
    [data-testid="stFormSubmitButton"] button,
    [data-testid="stLinkButton"] a {
       background:linear-gradient(135deg,#2479D1,#1557A7) !important;
       color:#FFFFFF !important;border:1px solid #1557A7 !important;
       border-radius:10px !important;font-size:1rem !important;
       font-weight:700 !important;min-height:43px !important;
       white-space:normal;box-shadow:0 3px 9px rgba(21,87,167,.12);
       transition:background .15s ease,box-shadow .15s ease,transform .15s ease;
    }
    [data-testid="stButton"] button :is(p,span),
    [data-testid="stDownloadButton"] button :is(p,span),
    [data-testid="stFormSubmitButton"] button :is(p,span),
    [data-testid="stLinkButton"] a :is(p,span) {
       color:#FFFFFF !important;font-size:1rem !important;font-weight:700 !important;
    }
    [data-testid="stButton"] button:hover,
    [data-testid="stDownloadButton"] button:hover,
    [data-testid="stFormSubmitButton"] button:hover,
    [data-testid="stLinkButton"] a:hover {
       background:linear-gradient(135deg,#1E67B8,#104889) !important;
       border-color:#104889 !important;
       box-shadow:0 6px 13px rgba(21,87,167,.23);
    }
    [data-testid="stButton"] button:focus-visible,
    [data-testid="stDownloadButton"] button:focus-visible,
    [data-testid="stFormSubmitButton"] button:focus-visible {
       outline:3px solid #91C7FB !important;outline-offset:2px;
    }
    [data-testid="stButton"] button:disabled,
    [data-testid="stDownloadButton"] button:disabled,
    [data-testid="stFormSubmitButton"] button:disabled {
       opacity:.52;cursor:not-allowed;box-shadow:none;
    }
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] button {
       width:100%;min-height:44px !important;
    }
    .ld-brand {display:flex;align-items:center;gap:.7rem;padding:.35rem .1rem 1rem}
    .ld-brand svg {flex:none;width:43px;height:43px}
    .ld-brand-name {font-weight:750;font-size:1.52rem;color:#FFFFFF;letter-spacing:-.025em}
    .ld-brand-name span {color:#8BC0FF}
    .ld-brand-caption {color:#C8E2FF;font-size:.70rem;letter-spacing:.09em}
    [data-testid="stMetric"] {background:#FFFFFF;border:1px solid #E4E8F2;border-radius:15px;
       min-width:0;min-height:132px;height:100%;padding:18px 16px 16px;
       box-shadow:0 7px 22px rgba(28,39,77,.045)}
    [data-testid="stMetricLabel"] {line-height:1.38;min-height:2.4em;white-space:normal;
       overflow-wrap:break-word;color:#60738F;font-size:clamp(.90rem,1.02vw,1.06rem);font-weight:650}
    [data-testid="stMetricValue"] {overflow:visible !important;min-width:0 !important;
       font-size:clamp(1.39rem,1.72vw,1.88rem) !important;font-weight:750;
       color:#202C4A;line-height:1.25;letter-spacing:-.027em}
    [data-testid="stMetricValue"]>div, [data-testid="stMetricValue"]>div>div,
    [data-testid="stMetricValue"] p {text-overflow:clip !important;overflow:visible !important;
       white-space:normal !important;overflow-wrap:anywhere !important;word-break:normal}
    [data-testid="stHorizontalBlock"] {align-items:stretch}
    [data-testid="column"] {min-width:0;overflow:visible}
    [data-baseweb="tab-list"] {gap:.25rem;flex-wrap:wrap}
    [data-baseweb="tab"] {border-radius:9px;padding:.55rem .77rem;font-weight:650;font-size:.98rem}
    [data-baseweb="tab"][aria-selected="true"] {color:#225BA9;background:#E9F2FD}
    [data-testid="stDataFrame"] {border:1px solid #E3E9F2;border-radius:13px}
    [data-testid="stPlotlyChart"] {background:#FFFFFF;border:1px solid #E7ECF3;
       border-radius:15px;padding:9px;box-shadow:0 6px 20px rgba(28,39,77,.03)}
    [data-testid="stMarkdownContainer"] p {line-height:1.48}
    [data-testid="stMarkdownContainer"] h3 {font-size:1.28rem !important}
    [data-testid="stMarkdownContainer"] h4 {font-size:1.13rem !important}
    .ld-hero {display:flex;align-items:center;gap:18px;padding:24px 27px;
       border:1px solid #E0E9F4;border-radius:16px;
       background:linear-gradient(115deg,#FFFFFF 22%,#EBF3FD 100%);
       box-shadow:0 8px 28px rgba(53,38,102,.05);margin-bottom:17px}
    .ld-hero-mark {width:45px;height:45px;border-radius:12px;background:#DFEDFD;
       display:flex;align-items:center;justify-content:center;flex:none}
    .ld-eyebrow {font-size:.76rem;font-weight:750;letter-spacing:.09em;color:#2865B8}
    .ld-title {font-size:clamp(1.50rem,1.95vw,2.02rem);font-weight:750;
       color:#1D294B;line-height:1.2;margin:.16rem 0 .20rem}
    .ld-description {font-size:.97rem;color:#66738A;margin:0}
    @media(max-width:960px) {[data-testid="stMetric"] {padding:13px;min-height:120px}
       [data-testid="stMetricValue"] {font-size:1.42rem !important}}
    @media(max-width:700px) {.ld-hero{padding:18px 16px;gap:12px}
       .ld-hero-mark{width:45px;height:45px}.ld-title{font-size:1.48rem}}
    </style>''', unsafe_allow_html=True)


    raw_dataframe = st.dataframe
    raw_plotly_chart = st.plotly_chart
    chart_serial = [0]
    # Solo etiquetas de presentación, sin alterar rutas ni identificadores originales.
    route_display_names = {}

    def _clean_route_key(value):
        raw = str(value).strip()
        for prefix in ('FII-', 'FI-'):
            if raw.upper().startswith(prefix):
                return raw[len(prefix):]
        return raw

    def _route_ordinal(value):
        raw = _clean_route_key(value)
        if raw in route_display_names:
            return route_display_names[raw]
        last = re.search(r'(?:R[-_])?(?:[A-Za-z0-9]+[-_])*0*(\d+)$', raw, re.I)
        return int(last.group(1)) if last else None

    def friendly_route(value):
        ordinal = _route_ordinal(value)
        return f'Ruta {ordinal}' if ordinal is not None else 'Ruta'

    def friendly_vehicle(value):
        ordinal = _route_ordinal(value)
        return f'Vehículo {ordinal}' if ordinal is not None else 'Vehículo'

    def friendly_alternative(value):
        raw = str(value).strip()
        if raw.upper().startswith('ADICIONAL-'):
            return 'Ciclo adicional ' + raw.split('-', 1)[1].lstrip('0')
        if re.fullmatch(r'\d+(?:\.0)?', raw):
            return f'Alternativa {int(float(raw))}'
        m = re.search(r'(\d+)$', raw)
        return f'Alternativa {int(m.group(1))}' if m else 'Alternativa'

    def user_facing_table(df):
        """Presentation copy: IDs stay unchanged in official Excel and CSV exports."""
        if not isinstance(df, pd.DataFrame):
            return df
        shown = df.copy()
        internal = {'solucion_id', 'id_solucion', 'solution_id', 'config_id',
                    'id_configuracion', 'configuracion_id', 'firma_solucion'}
        shown = shown.drop(columns=[c for c in shown if c.lower().strip() in internal], errors='ignore')
        for col in shown.columns:
            norm = str(col).lower().strip()
            if norm in ('ruta_id', 'ruta'):
                shown[col] = shown[col].map(friendly_route)
            elif norm == 'vehiculo':
                shown[col] = shown[col].map(friendly_vehicle)
            elif norm in ('candidato','candidata'):
                shown[col] = shown[col].map(friendly_alternative)
        return shown

    def data_table(df, **kwargs):
        # Dos decimales solo en el render: conservar los valores originales en memoria y exportación.
        if isinstance(df, pd.DataFrame):
            df = user_facing_table(df)
            numeric = [c for c in df.columns if pd.api.types.is_float_dtype(df[c])]
            if numeric and hasattr(st, 'column_config'):
                kwargs.setdefault('column_config', {c: st.column_config.NumberColumn(
                    str(c).replace('_',' '), format='%.2f') for c in numeric})
        kwargs.setdefault('use_container_width', True)
        kwargs.setdefault('hide_index', True)
        return raw_dataframe(df, **kwargs)

    def show_chart(fig, **kwargs):
        chart_serial[0] += 1
        is_map = any(str(tr.type) in ('scattermap', 'scattermapbox') for tr in fig.data)
        fig.update_layout(font=dict(family='Inter, Arial, sans-serif', size=14, color=NAVY),
            title_font=dict(size=18, color=NAVY),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=10,r=10,t=48,b=100 if is_map else 28),
            legend=dict(font=dict(size=12)), legend_title_text='',
            separators=',.')
        raw_plotly_chart(fig, **kwargs)
        with st.expander('Descargar esta gráfica (PNG / HTML)', expanded=False):
            download_fig(fig, f'LogiData_{source_name.rsplit(".",1)[0]}_grafica_{chart_serial[0]:02d}.html',
                         f'chart_{chart_serial[0]}')

    def download_table_xlsx(df, filename, key):
        if df is None or df.empty:
            return
        buffer = BytesIO()
        with pd.ExcelWriter(buffer,engine='openpyxl') as writer:
            df.to_excel(writer,sheet_name='Resultados',index=False)
        st.download_button('Descargar Excel',buffer.getvalue(),filename,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key=key)

    def num(v, default=np.nan):
        try:
            if v is None or pd.isna(v) or str(v).strip() == '':
                return default
            return float(v)
        except (ValueError, TypeError):
            return default


    def fmt_money(v):
        if not np.isfinite(num(v)):
            return 'No disponible'
        return ('$' + f'{float(v):,.2f}').replace(',', 'X').replace('.', ',').replace('X', '.')


    def fmt_num(v, decimals=1):
        if not np.isfinite(num(v)):
            return 'No disponible'
        return f'{float(v):,.{decimals}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


    def to_numeric(df, cols):
        for c in cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce')
        return df


    def first_existing(df, aliases, new, fallback=np.nan):
        for a in aliases:
            if a in df.columns:
                df[new] = df[a]
                return
        df[new] = fallback


    def clean_ids(df, cols):
        for col in cols:
            if col in df:
                df[col] = df[col].fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        return df


    @st.cache_data(show_spinner=False)
    def read_book(data: bytes):
        sheets = pd.read_excel(BytesIO(data), sheet_name=None, engine='openpyxl')
        return {name: df.dropna(how='all').copy() for name, df in sheets.items()}


    def sheet(book, name):
        for key in (name, name.upper(), name.lower()):
            if key in book:
                return book[key].copy()
        return pd.DataFrame()


    def phase_sheet(book, phase, suffix):
        return sheet(book, f'{phase}_{suffix}')


    def rowdict(book, name):
        df = sheet(book, name)
        return df.iloc[0].to_dict() if not df.empty else {}


    def pick(row, *keys, default=np.nan):
        for key in keys:
            if key in row and pd.notna(row[key]):
                return row[key]
        return default


    def build_phase(book, phase):
        """Normaliza SOLO nombres de columnas; mantiene valores tal como constan en Excel."""
        r = phase_sheet(book, phase, 'Rutas')
        v = phase_sheet(book, phase, 'Visitas')
        a = phase_sheet(book, phase, 'Asignaciones')
        b = phase_sheet(book, phase, 'BodegasAbiertas')
        s = phase_sheet(book, phase, 'Soluciones')
        clusters = phase_sheet(book, phase, 'Clusters')
        segments = phase_sheet(book, phase, 'Segmentos')
        pairings = phase_sheet(book, phase, 'Emparejamientos')
        members = phase_sheet(book, phase, 'MiembrosCluster')
        segment_visits = phase_sheet(book, phase, 'VisitasSegmento')
        if not r.empty:
            first_existing(r, ['ruta_id', 'id_ruta'], 'ruta_id')
            first_existing(r, ['bodega_id', 'deposito'], 'bodega_id')
            first_existing(r, ['km', 'distancia_km'], 'km')
            first_existing(r, ['minutos', 'tiempo_min'], 'minutos')
            first_existing(r, ['carga_inicial'], 'carga_inicial')
            first_existing(r, ['costo_transporte_diario', 'costo'], 'costo_transporte')
            first_existing(r, ['arriendo_diario_vehiculo', 'arriendo_vehiculo_diario'], 'arriendo_ruta')
            first_existing(r, ['costo_total_diario_ruta','costo_total_ruta_diario'], 'costo_diario_ruta')
            to_numeric(r, ['km','minutos','carga_inicial','costo_transporte','arriendo_ruta','costo_diario_ruta'])
            clean_ids(r, ['ruta_id','bodega_id'])
        if not v.empty:
            first_existing(v, ['ruta_id'], 'ruta_id')
            first_existing(v, ['nodo_id','cliente_id'], 'nodo_id')
            first_existing(v, ['operacion'], 'operacion')
            for key in ('orden','cantidad','carga_antes','carga_despues','km_arco','min_arco','km_acum','min_acum'):
                first_existing(v, [key], key)
            to_numeric(v, ['orden','cantidad','carga_antes','carga_despues','km_arco','min_arco','km_acum','min_acum'])
            clean_ids(v,['ruta_id','nodo_id'])
            v['operacion'] = v['operacion'].astype(str).str.strip().str.upper()
            v = v.sort_values(['ruta_id','orden'])
        if not a.empty:
            first_existing(a, ['cliente_id'], 'cliente_id')
            first_existing(a, ['bodega_id'], 'bodega_id')
            first_existing(a, ['entrega'], 'entrega')
            first_existing(a, ['recogida'], 'recogida')
            to_numeric(a,['entrega','recogida'])
            clean_ids(a,['cliente_id','bodega_id'])
        if not b.empty:
            aliases = {
                'bodega_id':['bodega_id'], 'capacidad':['capacidad'],
                'costo_diario_bodega':['costo_diario_bodega','costo_apertura'],
                'entrega_asignada':['entrega_asignada'], 'recogida_asignada':['recogida_asignada'],
                'vehiculos_utilizados':['vehiculos_utilizados','vehiculos_usados'],
                'vehiculos_salen':['vehiculos_salen'], 'vehiculos_regresan':['vehiculos_regresan'],
                'costo_arriendo_diario':['costo_arriendo_diario','arriendo_vehiculos_diario'],
                'km_diarios':['km_diarios'], 'costo_transporte_diario':['costo_transporte_diario'],
            }
            for k, vals in aliases.items():
                first_existing(b, vals, k)
            to_numeric(b, list(aliases.keys())[1:])
            clean_ids(b,['bodega_id'])
        if not s.empty:
            for k, aliases in {'costo_total':['costo_total'], 'costo_apertura':['costo_apertura'],
                  'costo_arriendo':['costo_arriendo_vehiculos'], 'costo_transporte':['costo_transporte'],
                  'km':['km'], 'rutas':['rutas']}.items():
                first_existing(s, aliases, k)
            to_numeric(s, ['costo_total','costo_apertura','costo_arriendo','costo_transporte','km','rutas'])
        for df in (clusters, segments):
            if not df.empty and 'bodega_id' in df:
                clean_ids(df,['bodega_id'])
        return dict(phase=phase, routes=r, visits=v, assignments=a, depots=b,
                    solutions=s, clusters=clusters, segments=segments, pairings=pairings,
                    members=members, segment_visits=segment_visits,
                    detailed=not r.empty and not v.empty)


    def money_metric(label, value, delta=None):
        st.metric(label, fmt_money(value), delta=delta, delta_color='inverse')


    def phase_summary(phase_data, comparison, phase):
        is_fi = phase == 'FI'
        suffix = 'fase1' if is_fi else 'fase2'
        sol = phase_data['solutions']
        obj = sol.iloc[0].to_dict() if not sol.empty else {}
        def col(*names, default=np.nan):
            return pick(obj, *names, default=default)
        total = pick(comparison, 'costo_inicial' if is_fi else 'costo_final', default=col('costo_total'))
        fixed = pick(comparison, f'costo_diario_bodegas_{suffix}', f'bodegas_diario_{suffix}', default=col('costo_apertura'))
        rent = pick(comparison, f'costo_diario_arriendo_{suffix}', f'arriendo_diario_{suffix}', default=col('costo_arriendo'))
        variable = pick(comparison, f'costo_diario_transporte_{suffix}', f'transporte_diario_{suffix}', default=col('costo_transporte'))
        km = pick(comparison, f'km_{suffix}', default=col('km'))
        fleet = pick(comparison, f'vehiculos_{suffix}', f'rutas_{suffix}', default=col('rutas'))
        depots_str = pick(comparison, f'bodegas_{suffix}', default='')
        if phase_data['detailed']:
            v = phase_data['visits']
            aa = phase_data['assignments']
            total_minutes = phase_data['routes']['minutos'].sum(min_count=1)
            clients = aa['cliente_id'].nunique() if not aa.empty else v.loc[v.operacion.isin(['DELIVERY','PICKUP']),'nodo_id'].nunique()
            ops = int(v.operacion.isin(['DELIVERY','PICKUP']).sum())
            depots_count = phase_data['depots']['bodega_id'].nunique() if not phase_data['depots'].empty else phase_data['routes']['bodega_id'].nunique()
            fleet = len(phase_data['routes'])
            km = phase_data['routes']['km'].sum()
        else:
            total_minutes = np.nan
            clients = np.nan
            ops = np.nan
            depots_count = len([x for x in re.split(r'\s*\|\s*',str(depots_str)) if x]) if depots_str else np.nan
        return {'fase':phase, 'costo_total':num(total), 'bodegas_cost':num(fixed),
                'arriendo':num(rent), 'transporte':num(variable), 'km':num(km),
                'rutas':num(fleet), 'vehiculos':num(fleet), 'bodegas':depots_count,
                'minutos':num(total_minutes), 'clientes':num(clients), 'operaciones':num(ops),
                'bodegas_ids':depots_str}


    def coords_table(df, kind):
        if df.empty:
            return pd.DataFrame(columns=['nodo_id','lat','lon'])
        low = {str(c).lower().strip():c for c in df.columns}
        names = (('cliente', 'id', 'cliente_id') if kind == 'cliente' else ('id','bodega_id'))
        idcol = next((low[k] for k in names if k in low), None)
        latcol = next((low[k] for k in ('latitud','lat','latitude','y') if k in low), None)
        loncol = next((low[k] for k in ('longitud','lon','longitude','lng','x') if k in low), None)
        if not (idcol and latcol and loncol):
            return pd.DataFrame(columns=['nodo_id','lat','lon'])
        out = df[[idcol, latcol, loncol]].copy()
        out.columns = ['nodo_id','lat','lon']
        out['nodo_id'] = out.nodo_id.astype(str).str.strip()
        to_numeric(out,['lat','lon'])
        out = out[(out.lat.between(-90,90)) & (out.lon.between(-180,180))]
        return out.drop_duplicates('nodo_id')


    @st.cache_data(show_spinner=False)
    def read_coords(data, kind):
        df = pd.read_excel(BytesIO(data), sheet_name=None, engine='openpyxl')
        return coords_table(next(iter(df.values())), kind)


    def map_figure(routes, visits, coords, title='Rutas LogiData', modo_coordenadas=False, map_order_route=None):
        """Map of actual recorded visit order; connecting lines are schematic."""
        if coords.empty or routes.empty or visits.empty:
            return None
        coords0 = coords.copy()
        coords0['nodo_id'] = coords0.nodo_id.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        lookup = coords0.drop_duplicates('nodo_id').set_index('nodo_id')[['lat','lon']].to_dict('index')
        fig = go.Figure()
        used_positions = []
        displayed_stops = []
        no_coords = []
        one_route = len(routes) == 1
        depots_in_selection = []
        # Lista de marcas numeradas: se dibujan encima de los clientes al final.
        number_traces = []
        for idx, route in routes.reset_index(drop=True).iterrows():
            depot_id = str(route.bodega_id)
            depots_in_selection.append(depot_id)
            stops = visits[visits.ruta_id == route.ruta_id].sort_values('orden')
            if stops.empty:
                continue
            # Show the warehouse at both ends even when it is recorded as two stops.
            ids = stops.nodo_id.astype(str).str.strip().tolist()
            operations = stops.operacion.astype(str).str.upper().tolist()
            if operations[0] != 'DEPOT_START' or operations[-1] != 'DEPOT_END':
                # Never fabricate a route when the exported sequence is incomplete.
                no_coords.append(friendly_route(route.ruta_id))
                continue
            # Use precisely the recorded warehouse/visits and the recorded order.
            color = PALETTE[idx % len(PALETTE)]
            segments_x, segments_y = [], []
            valid_count = 0
            for node in ids:
                position = lookup.get(node)
                if position is None:
                    segments_x.append(None);segments_y.append(None)
                    continue
                segments_x.append(position['lon'])
                segments_y.append(position['lat'])
                used_positions.append(position)
                valid_count += 1
            if valid_count < 2:
                no_coords.append(friendly_route(route.ruta_id));continue
            if valid_count != len(ids):
                no_coords.append(friendly_route(route.ruta_id))
            # Do not draw fake shortcuts across stops without coordinates.
            fig.add_trace(go.Scattermap(
                lat=segments_y, lon=segments_x, mode='lines',
                line=dict(color=color, width=5 if one_route else 3.5),
                name=friendly_vehicle(route.ruta_id),
                customdata=[[str(route.ruta_id),num(route.km,0),num(route.minutos,0)]
                            if x is not None else [str(route.ruta_id),0,0] for x in segments_x],
                hovertemplate='Recorrido · %{customdata[1]:.2f} km · %{customdata[2]:.2f} min<extra></extra>',
                showlegend=True))
            displayed_stops.append(stops)
            # Numeración del recorrido: siempre en la ruta individual y, en el
            # mapa general, en la ruta elegida para evitar superposiciones.
            wants_order = one_route or (map_order_route is not None and str(map_order_route) == str(route.ruta_id))
            if wants_order:
                numbered = stops[stops.operacion.isin(['DELIVERY','PICKUP'])].copy()
                numbered = numbered[numbered.nodo_id.astype(str).isin(lookup)]
                if not numbered.empty:
                    labels, locations, hover_details, bubble_sizes = [], [], [], []
                    # Dos operaciones en el mismo cliente comparten coordenadas.
                    # Mostrar, por ejemplo, «2/5» sobre un único marcador evita
                    # ocultar el orden real y no altera el trazado geométrico.
                    for node, visits_at_node in numbered.groupby('nodo_id', sort=False):
                        orders = [str(int(z)) for z in visits_at_node.orden]
                        labels.append('/'.join(orders))
                        locations.append(lookup[str(node)])
                        bubble_sizes.append(25 if len(orders)>1 else 23)
                        detail = []
                        for _, stop in visits_at_node.iterrows():
                            task = 'Entrega' if stop.operacion == 'DELIVERY' else 'Recogida'
                            detail.append(f'Paso {int(stop.orden)} · {task}: '
                                f'{fmt_num(stop.cantidad,2)} unidades · '
                                f'carga posterior {fmt_num(stop.carga_despues,2)} unidades')
                        hover_details.append(f'Cliente {node}<br>'+'<br>'.join(detail)+
                                             f'<br>{friendly_vehicle(route.ruta_id)} · Bodega {depot_id}')
                    number_traces.append(go.Scattermap(
                        lat=[x['lat'] for x in locations],
                        lon=[x['lon'] for x in locations],
                        mode='markers+text',
                        marker=dict(size=bubble_sizes,color='#FFFFFF',opacity=.97),
                        text=labels,
                        textposition='middle center',textfont=dict(size=12,color=NAVY),
                        customdata=hover_details,
                        hovertemplate='%{customdata}<extra></extra>',
                        name=f'Orden de visita · {friendly_vehicle(route.ruta_id)}',
                        showlegend=True))
        if not displayed_stops:
            return None
        stops_all = pd.concat(displayed_stops, ignore_index=True)
        client_stops = stops_all[stops_all.operacion.isin(['DELIVERY','PICKUP'])]
        client_types = client_stops.groupby('nodo_id').operacion.agg(lambda x:set(x)).to_dict()
        for kind, color, predicate in (
            ('Mixto', VIOLET, lambda x: client_types.get(x,set()) == {'DELIVERY','PICKUP'}),
            ('Entrega', BLUE, lambda x: client_types.get(x,set()) == {'DELIVERY'}),
            ('Recogida', ORANGE, lambda x: client_types.get(x,set()) == {'PICKUP'})):
            subset = client_stops[client_stops.nodo_id.map(predicate)].drop_duplicates('nodo_id')
            subset = subset[subset.nodo_id.astype(str).isin(lookup)]
            if subset.empty:
                continue
            texts = []
            for node in subset.nodo_id:
                at_node = client_stops[client_stops.nodo_id == node]
                detail = '<br>'.join(f'Visita {int(v.orden)} · {"Entrega" if v.operacion=="DELIVERY" else "Recogida"}: {float(v.cantidad):.2f} u'
                                   for _, v in at_node.iterrows())
                texts.append(f'Cliente {node} · {kind}<br>{detail}')
            fig.add_trace(go.Scattermap(
                lat=[lookup[str(n)]['lat'] for n in subset.nodo_id],
                lon=[lookup[str(n)]['lon'] for n in subset.nodo_id],
                mode='markers', marker=dict(size=12 if one_route else 10, color=color),
                text=texts, hovertemplate='%{text}<extra></extra>',
                name=f'Clientes · {kind}', showlegend=True))
        depot_ids = list(dict.fromkeys(d for d in depots_in_selection if d in lookup))
        if depot_ids:
            # Marker above all route lines: warehouses are large dark blue markers,
            # always contrasted against smaller colored customer circles.
            fig.add_trace(go.Scattermap(
                lat=[lookup[n]['lat'] for n in depot_ids],
                lon=[lookup[n]['lon'] for n in depot_ids],
                mode='markers+text', marker=dict(size=27,color='#1854A2'),
                text=depot_ids, textfont=dict(size=11,color='#FFFFFF'),
                textposition='middle center',
                customdata=depot_ids,
                hovertemplate='Bodega %{customdata}<br>Inicio y retorno de rutas<extra></extra>',
                name='Bodegas · salida y regreso',showlegend=True))
        for number_trace in number_traces:
            fig.add_trace(number_trace)
        if not used_positions:
            return None
        lat_values = [p['lat'] for p in used_positions]
        lon_values = [p['lon'] for p in used_positions]
        span = max(max(lat_values)-min(lat_values),max(lon_values)-min(lon_values))
        zoom = 12 if span<.035 else 11 if span<.075 else 10 if span<.17 else 9 if span<.37 else 8
        fig.update_layout(title=title, map=dict(style='open-street-map',
            center=dict(lat=(min(lat_values)+max(lat_values))/2,
                        lon=(min(lon_values)+max(lon_values))/2), zoom=zoom),
            height=610, margin=dict(l=5,r=5,t=48,b=105),
            legend=dict(orientation='h',y=-.12,x=0,font=dict(size=10)),
            paper_bgcolor='rgba(0,0,0,0)')
        if no_coords:
            fig.add_annotation(text=f'{len(set(no_coords))} ruta(s) con alguna coordenada ausente; no se inventan tramos',
                xref='paper',yref='paper',x=0,y=0,showarrow=False,font=dict(color=RED,size=10))
        if modo_coordenadas:
            # Without internet access: still show every verified geographic point,
            # warehouse, and route segment on longitude/latitude axes.
            offline=go.Figure()
            for tr in fig.data:
                original_mode=str(getattr(tr,'mode','markers'))
                args={'x':tr.lon,'y':tr.lat, 'mode':original_mode,
                      'name':tr.name,'showlegend':tr.showlegend,
                      'hovertemplate':getattr(tr,'hovertemplate',None)}
                if 'lines' in original_mode:
                    args['line']=dict(color=tr.line.color,width=tr.line.width)
                if 'markers' in original_mode:
                    args['marker']=dict(size=tr.marker.size,color=tr.marker.color)
                if 'text' in original_mode:
                    args.update(text=tr.text,textposition=tr.textposition,
                                textfont=dict(size=11,color=NAVY))
                if getattr(tr,'customdata',None) is not None:
                    args['customdata']=tr.customdata
                offline.add_trace(go.Scatter(**args))
            offline.update_layout(title=title,height=600,margin=dict(l=55,r=20,t=48,b=90),
                xaxis=dict(title='Longitud (°)',tickformat='.3f',showgrid=True,gridcolor='#E4ECF5'),
                yaxis=dict(title='Latitud (°)',tickformat='.3f',showgrid=True,
                           gridcolor='#E4ECF5',scaleanchor='x',scaleratio=1),
                legend=dict(orientation='h',y=-.15,x=0,font=dict(size=10)),
                plot_bgcolor='#F9FCFF',paper_bgcolor='#FFFFFF')
            return offline
        return fig

    def route_table(pdata, capacity=2500):
        r, v = pdata['routes'], pdata['visits']
        if r.empty or v.empty:
            return pd.DataFrame()
        rows=[]
        for _, route in r.iterrows():
            stops=v[v.ruta_id==route.ruta_id].sort_values('orden')
            real=stops[stops.operacion.isin(['DELIVERY','PICKUP'])]
            start=stops.iloc[0] if not stops.empty else None
            end=stops.iloc[-1] if not stops.empty else None
            rows.append(dict(fase=pdata['phase'],ruta=route.ruta_id,vehiculo=f"{pdata['phase']}-{route.ruta_id}",
                             bodega=route.bodega_id,origen=start.nodo_id if start is not None else None,
                             retorno=end.nodo_id if end is not None else None,
                             visitas=len(real),clientes=real.nodo_id.nunique(),
                             carga_inicial=route.carga_inicial,
                             entregas=int((real.operacion=='DELIVERY').sum()),
                             recogidas=int((real.operacion=='PICKUP').sum()),
                             unidades_entregadas=real.loc[real.operacion=='DELIVERY','cantidad'].sum(),
                             unidades_recogidas=real.loc[real.operacion=='PICKUP','cantidad'].sum(),
                             secuencia=' → '.join(stops.nodo_id.astype(str)),
                             km=route.km,minutos=route.minutos,
                             pico_carga=stops['carga_despues'].max(),
                             utilizacion_capacidad_pct=100*stops['carga_despues'].max()/capacity if capacity>0 else np.nan,
                             costo_transporte=route.costo_transporte,
                             arriendo=route.arriendo_ruta,
                             costo_ruta_diario=route.costo_diario_ruta))
        return pd.DataFrame(rows)


    def ordered_visit_table(stops, depot_id):
        """Vista ordenada de la ruta; no modifica las operaciones exportadas."""
        if stops.empty:
            return pd.DataFrame()
        pretty = stops.sort_values('orden').copy()
        actions = {'DEPOT_START':'Salida de bodega', 'DEPOT_END':'Regreso a bodega',
                   'DELIVERY':'Entrega', 'PICKUP':'Recogida'}
        pretty['Paso'] = pd.to_numeric(pretty['orden'],errors='coerce').fillna(0).astype(int)
        pretty['Lugar'] = pretty.apply(lambda r: ('Bodega ' if r['operacion'] in ('DEPOT_START','DEPOT_END')
                     else 'Cliente ') + str(r['nodo_id']),axis=1)
        pretty['Operación'] = pretty['operacion'].map(actions).fillna(pretty['operacion'])
        pretty = pretty.rename(columns={'cantidad':'Unidades','km_arco':'Km desde anterior',
            'min_arco':'Minutos desde anterior','km_acum':'Km acumulados',
            'min_acum':'Minutos acumulados','carga_antes':'Carga antes',
            'carga_despues':'Carga después'})
        return pretty[['Paso','Lugar','Operación','Unidades','Km desde anterior',
                       'Minutos desde anterior','Km acumulados','Minutos acumulados',
                       'Carga antes','Carga después']].reset_index(drop=True)

    def readable_sequence(stops, max_chars=180):
        if stops.empty:
            return ''
        seq=[]
        for _,p in stops.sort_values('orden').iterrows():
            code=str(p['operacion']).upper()
            if code=='DEPOT_START':
                seq.append('Salida Bodega '+str(p['nodo_id']))
            elif code=='DEPOT_END':
                seq.append('Regreso Bodega '+str(p['nodo_id']))
            else:
                seq.append(str(int(p['orden']))+'. '+str(p['nodo_id'])+
                           (' (entrega)' if code=='DELIVERY' else ' (recogida)'))
        rendered='  →  '.join(seq)
        return rendered if len(rendered)<=max_chars else rendered[:max_chars].rsplit('  →  ',1)[0]+'  →  …'

    def savings_commentary(stage_i, stage_ii, label, unit='COP/día'):
        a,b=num(stage_i),num(stage_ii)
        if not (np.isfinite(a) and np.isfinite(b)):
            return f'{label}: los datos de ambas fases no están completos para calcular el cambio.'
        diff=a-b
        if abs(diff)<1e-7:
            return f'{label}: no cambió entre la Fase I y la Fase II ({fmt_num(a,2)} {unit}).'
        direction='disminuyó' if diff>0 else 'aumentó'
        pct=f' ({fmt_num(abs(diff)/abs(a)*100,2)} %)' if abs(a)>1e-9 else ''
        return (f'{label} {direction} en {fmt_num(abs(diff),2)} {unit}{pct}: '
                f'de {fmt_num(a,2)} a {fmt_num(b,2)} {unit}.')

    def verify_phase(pdata, cost_per_km=None, vehicle_capacity=2500):
        """Auditoría independiente a partir de hojas exportadas, nunca asume OK."""
        checks=[]
        r,v,a,b,s=(pdata[k] for k in ('routes','visits','assignments','depots','solutions'))
        def add(name, result, detail, available=True):
            checks.append({'Fase':pdata['phase'],'Restricción / control':name,
                           'Resultado':'No verificable' if not available else ('Cumple' if result else 'Incumple'),
                           'Detalle':detail})
        if r.empty or v.empty:
            add('Integridad de rutas y visitas',False,'No se exportaron estructuras de esta fase',False)
            if not s.empty:
                add('Validación registrada por el algoritmo',str(s.iloc[0].get('validacion','')).upper()=='OK',
                    'Indicador exportado; no reemplaza la auditoría independiente',True)
            return pd.DataFrame(checks)
        pairs_ok=True
        departures=Counter(); returns=Counter(); seen={}
        load_ok=True; prec_ok=True; arcs_ok=True; route_cost_ok=True
        for _, route in r.iterrows():
            vv=v[v.ruta_id==route.ruta_id].sort_values('orden')
            if len(vv)<3:
                pairs_ok=False; arcs_ok=False
                continue
            start,end=vv.iloc[0],vv.iloc[-1]
            valid_start=start.operacion=='DEPOT_START' and start.nodo_id==route.bodega_id
            valid_end=end.operacion=='DEPOT_END' and end.nodo_id==route.bodega_id
            if valid_start: departures[route.bodega_id]+=1
            if valid_end: returns[route.bodega_id]+=1
            pairs_ok &= valid_start and valid_end
            actual=vv.iloc[1:-1]
            seq_ops=actual.operacion.tolist()
            prec_ok &= all(t=='DELIVERY' for t in seq_ops[:seq_ops.index('PICKUP')]) if 'PICKUP' in seq_ops else True
            prec_ok &= not any(t=='DELIVERY' for t in seq_ops[seq_ops.index('PICKUP')+1:]) if 'PICKUP' in seq_ops else True
            prec_ok &= all(t in ('DELIVERY','PICKUP') for t in seq_ops)
            load_ok &= bool((vv.carga_antes >= -0.01).all() and (vv.carga_despues >= -0.01).all()
                            and (vv.carga_despues <= vehicle_capacity+0.01).all())
            load_ok &= bool((vv['carga_despues'].iloc[1:].to_numpy() - vv['carga_antes'].iloc[1:].to_numpy()
                             - np.where(vv['operacion'].iloc[1:].eq('DELIVERY'), -vv['cantidad'].iloc[1:],
                              np.where(vv['operacion'].iloc[1:].eq('PICKUP'), vv['cantidad'].iloc[1:],0))).__abs__().max()<.01)
            arcs_ok &= (list(vv['orden']) == list(range(len(vv))))
            arcs_ok &= abs(vv.km_arco.sum()-num(route.km,0))<.02
            arcs_ok &= abs(vv.min_arco.sum()-num(route.minutos,0))<.02
            if np.isfinite(num(cost_per_km)):
                route_cost_ok &= abs(num(route.costo_transporte,0)-num(route.km,0)*cost_per_km)<.03
            for _,visit in actual.iterrows():
                key=(str(visit.nodo_id),str(visit.operacion))
                seen[key]=seen.get(key,0)+1
        add('Salida y retorno a la misma bodega',pairs_ok, f'{len(r)} rutas revisadas')
        add('Vehículos = rutas = salidas = retornos',pairs_ok and departures==returns and sum(departures.values())==len(r),
            f'Rutas {len(r)} · Salidas {sum(departures.values())} · Retornos {sum(returns.values())}')
        add('Precedencia de entregas antes de recogidas',prec_ok,'Secuencia de operaciones inspeccionada')
        add('Capacidad vehicular y balances de carga',load_ok, f'Capacidad {vehicle_capacity:,.0f} unidades')
        add('Continuidad de la secuencia y arcos contabilizados',arcs_ok,
            'Se verifica secuencia explícita depósito-clientes-depósito; no geometrías viales')
        add('Costo variable del recorrido',route_cost_ok,
            'Costo reportado frente a km × tarifa' if np.isfinite(num(cost_per_km)) else 'Tarifa por km ausente',
            np.isfinite(num(cost_per_km)))
        if a.empty:
            for check in ('Cobertura exacta de entregas/recogidas','Clientes mixtos en una instalación',
                          'Capacidad individual de bodegas'):
                add(check,False,'No hay hoja de asignaciones',False)
        else:
            expected={}
            for _,row in a.iterrows():
                if num(row.entrega,0)>0: expected[(str(row.cliente_id),'DELIVERY')]=1
                if num(row.recogida,0)>0: expected[(str(row.cliente_id),'PICKUP')]=1
            coverage=expected==seen
            add('Cobertura exacta de entregas y recogidas',coverage,
                f'Esperadas {sum(expected.values())} · Realizadas {sum(seen.values())}')
            mixed=a[(a.entrega>0)&(a.recogida>0)]
            mixed_ok=bool(a.groupby('cliente_id').bodega_id.nunique().le(1).all())
            lookup=a.drop_duplicates('cliente_id').set_index('cliente_id').bodega_id.to_dict()
            for _,route in r.iterrows():
                for c in v[(v.ruta_id==route.ruta_id)&(v.operacion.isin(['DELIVERY','PICKUP']))].nodo_id:
                    mixed_ok &= (lookup.get(c)==route.bodega_id)
            add('Asignación única de clientes (incluidos mixtos)',mixed_ok,
                f'{len(mixed)} clientes mixtos contrastados con rutas')
            if b.empty:
                add('Capacidad individual de bodegas',False,'No hay capacidades exportadas',False)
            else:
                assigned=a.groupby('bodega_id').entrega.sum()
                cap_ok=all(assigned.get(row.bodega_id,0)<=row.capacidad+0.01 for _,row in b.iterrows())
                add('Capacidad individual de bodegas',cap_ok,'Capacidad calculada con entregas completas')
        if not s.empty and not b.empty:
            official=s.iloc[0]
            fixed=b.costo_diario_bodega.sum()
            rent=r.arriendo_ruta.sum(min_count=1)
            transport=r.costo_transporte.sum()
            match= all(abs(num(official[key],0)-val) < .03 for key,val in
                       [('costo_apertura',fixed),('costo_arriendo',rent),('costo_transporte',transport),
                        ('costo_total',fixed+rent+transport)])
            add('Reconciliación de función objetivo diaria',bool(match),
                f'Bodegas {fmt_money(fixed)} + arriendo {fmt_money(rent)} + transporte {fmt_money(transport)}')
        else:
            add('Reconciliación de función objetivo diaria',False,'Falta hoja de costos de la fase',False)
        add('Ausencia de subtours desconectados',pairs_ok and arcs_ok,
            'Verificable sobre la representación explícita como lista única de visitas, no sobre geometría vial')
        return pd.DataFrame(checks)


    def download_csv(df, filename, key):
        if not df.empty:
            st.download_button('Exportar tabla CSV', df.to_csv(index=False).encode('utf-8-sig'),
                               filename, mime='text/csv', key=key)


    def _png_de_respaldo(fig):
        """PNG local sin Chrome: dibuja los datos del gráfico, nunca cifras inventadas.

        Cuando no hay Kaleido se usa matplotlib. Para mapas conserva las
        coordenadas y polilíneas reales disponibles, sin callejero.
        """
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        output=BytesIO()
        fig_m,ax=plt.subplots(figsize=(10,5.6))
        plotted=False
        for series in fig.data:
            kind=str(series.type)
            label=str(series.name or kind)
            xs=getattr(series,'x',None)
            ys=getattr(series,'y',None)
            if kind in ('scattermapbox','scattermap'):
                xs=getattr(series,'lon',None)
                ys=getattr(series,'lat',None)
                if xs is None or ys is None: continue
                x=list(xs);y=list(ys)
                if not x or not y: continue
                if 'lines' in str(getattr(series,'mode','')):
                    ax.plot(x,y,linewidth=2,label=label)
                else:
                    ax.scatter(x,y,s=23,label=label)
                ax.set_xlabel('Longitud');ax.set_ylabel('Latitud')
                plotted=True
            elif kind=='waterfall':
                vals=list(getattr(series,'y',[]) if getattr(series,'y',None) is not None else [])
                labs=list(getattr(series,'x',[]) if getattr(series,'x',None) is not None else [])
                if vals and labs:
                    running=0.0
                    for i,(name,val) in enumerate(zip(labs,vals)):
                        value=float(val)
                        if i==0:base=0.0;running=value
                        elif i==len(vals)-1:base=0.0;value=running
                        else:base=running;running+=value
                        ax.bar(str(name),value,bottom=base,color=TEAL if value<0 else NAVY)
                    plotted=True
            elif kind=='bar' and xs is not None and ys is not None:
                ax.bar([str(x) for x in xs],list(ys),alpha=.8,label=label)
                plotted=True
            elif kind=='scatter' and xs is not None and ys is not None:
                ax.plot(list(xs),list(ys),marker='o' if 'markers' in str(getattr(series,'mode','')) else None,
                        label=label)
                plotted=True
        if not plotted:
            plt.close(fig_m)
            raise ValueError('Gráfico no soportado para la exportación PNG de respaldo')
        title=getattr(getattr(fig.layout,'title',None),'text',None)
        ax.set_title(str(title or 'LogiData'))
        ax.grid(alpha=.15)
        ax.legend(loc='best',fontsize=8)
        fig_m.autofmt_xdate(rotation=20)
        fig_m.tight_layout()
        fig_m.savefig(output,format='png',dpi=160)
        plt.close(fig_m)
        return output.getvalue()


    def download_fig(fig, filename, key):
        if fig is None: return
        st.download_button('Exportar gráfico HTML interactivo',
              fig.to_html(include_plotlyjs=True,full_html=True).encode('utf-8'),
              filename,mime='text/html',key=key)
        if st.button('Preparar PNG',key=key+'_prepare_png'):
            try:
                st.session_state[key+'_png_bytes']=fig.to_image(format='png',width=1300,height=730,scale=1.5)
            except Exception:
                try:
                    st.session_state[key+'_png_bytes']=_png_de_respaldo(fig)
                except Exception as e:
                    st.warning(f'El gráfico no pudo exportarse en PNG: {e}')
        if key+'_png_bytes' in st.session_state:
            st.download_button('Descargar PNG',st.session_state[key+'_png_bytes'],
               filename.rsplit('.',1)[0]+'.png',mime='image/png',key=key+'_download_png')


    def bar_compare(data, keys, title, unit):
        entries=[]
        for phase in ('FI','FII'):
            for field,label in keys:
                val=data[phase][field]
                if np.isfinite(val):
                    entries.append({'Fase':'Fase I' if phase=='FI' else 'Fase II',
                                    'Indicador':label,'Valor':val})
        if not entries:
            return None
        fig=px.bar(pd.DataFrame(entries),x='Indicador',y='Valor',color='Fase',barmode='group',
                   color_discrete_map={'Fase I':BLUE,'Fase II':TEAL},text_auto='.2f',title=title)
        fig.update_layout(yaxis_title=unit,xaxis_title='',legend_title='',height=370,
                          margin=dict(l=8,r=8,t=55,b=15))
        return fig


    st.sidebar.markdown('''<div class="ld-brand"><svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" fill="none" aria-label="Camión azul de LogiData">
<rect x="2" y="2" width="44" height="44" rx="12" fill="#E8F3FF"/>
<path d="M9 18h21v16H9z" rx="2" fill="#267BD1"/>
<path d="M30 23h5.6l6 7v4H30z" fill="#14549E"/>
<path d="M33 25.5h2l3.3 4H33z" fill="#D4EBFF"/>
<path d="M6 22h5M4 27h7" stroke="#68A5E4" stroke-width="2" stroke-linecap="round"/>
<circle cx="16" cy="35" r="4" fill="#123D76" stroke="#E8F3FF" stroke-width="1.5"/>
<circle cx="35" cy="35" r="4" fill="#123D76" stroke="#E8F3FF" stroke-width="1.5"/>
<circle cx="16" cy="35" r="1.4" fill="#D6EAFE"/>
<circle cx="35" cy="35" r="1.4" fill="#D6EAFE"/>
</svg>
       <div><div class="ld-brand-name">Logi<span>Data</span></div>
       <div class="ld-brand-caption">BOGOTÁ · LOGÍSTICA</div></div></div>''',
       unsafe_allow_html=True)
    arg_dir = None
    if '--results-dir' in sys.argv:
        i = sys.argv.index('--results-dir')
        if i + 1 < len(sys.argv):
            arg_dir = sys.argv[i + 1]
    resultados_dir = Path(arg_dir).expanduser().resolve() if arg_dir else Path(__file__).resolve().parent / 'resultados_logidata'
    archivos_instancias = sorted(resultados_dir.glob('LogiData_Resultados_*_clientes.xlsx')) if resultados_dir.is_dir() else []
    cli_excel = None
    if '--excel' in sys.argv:
        i = sys.argv.index('--excel')
        if i + 1 < len(sys.argv):
            cli_excel = Path(sys.argv[i+1]).expanduser().resolve()
    catalogo = list(archivos_instancias)
    if cli_excel is not None and cli_excel.is_file() and cli_excel not in catalogo:
        catalogo.insert(0,cli_excel)
    if not catalogo:
        fallback = Path(__file__).resolve().parent / 'LogiData_Fases_I_II_consolidado.xlsx'
        if fallback.is_file(): catalogo=[fallback]
    archivo_seleccionado = st.sidebar.selectbox('Instancia / ejecución',catalogo,
        index=catalogo.index(cli_excel) if cli_excel in catalogo else 0,
        format_func=lambda p: p.stem.replace('LogiData_Resultados_','').replace('_',' ')) if catalogo else None
    source_bytes=archivo_seleccionado.read_bytes() if archivo_seleccionado is not None else None
    source_name=archivo_seleccionado.name if archivo_seleccionado is not None else ''
    st.markdown('''<div class="ld-hero"><div class="ld-hero-mark"><svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" fill="none" aria-label="Camión azul de LogiData">
<rect x="2" y="2" width="44" height="44" rx="12" fill="#E8F3FF"/>
<path d="M9 18h21v16H9z" rx="2" fill="#267BD1"/>
<path d="M30 23h5.6l6 7v4H30z" fill="#14549E"/>
<path d="M33 25.5h2l3.3 4H33z" fill="#D4EBFF"/>
<path d="M6 22h5M4 27h7" stroke="#68A5E4" stroke-width="2" stroke-linecap="round"/>
<circle cx="16" cy="35" r="4" fill="#123D76" stroke="#E8F3FF" stroke-width="1.5"/>
<circle cx="35" cy="35" r="4" fill="#123D76" stroke="#E8F3FF" stroke-width="1.5"/>
<circle cx="16" cy="35" r="1.4" fill="#D6EAFE"/>
<circle cx="35" cy="35" r="1.4" fill="#D6EAFE"/>
</svg></div>
      <div><div class="ld-eyebrow">LOGIDATA · ANÁLISIS ACADÉMICO</div>
      <div class="ld-title">Resultados de localización y ruteo</div>
      <p class="ld-description">Bodegas, rutas, costos y ahorros de cada instancia · Bogotá</p></div></div>''',
      unsafe_allow_html=True)
    if source_bytes is None:
        st.info('No hay ejecuciones guardadas. Ejecuta una instancia desde Python para visualizar sus resultados.')
        st.stop()
    try:
        book=read_book(source_bytes)
    except Exception as exc:
        st.error(f'No fue posible leer el Excel: {exc}')
        st.stop()

    comparison=rowdict(book,'Comparacion_Fases')
    control=rowdict(book,'Control_Ejecucion')
    phase_data={p:build_phase(book,p) for p in ('FI','FII')}
    # Numeración legible y unívoca entre bodegas, independiente del ID interno.
    all_route_ids=sorted({str(x) for phase in phase_data.values()
        for x in (phase['routes']['ruta_id'].tolist() if not phase['routes'].empty else [])})
    route_display_names.update({route_id:i+1 for i,route_id in enumerate(all_route_ids)})
    summaries={p:phase_summary(phase_data[p],comparison,p) for p in ('FI','FII')}
    raw_fee=pick(control,'costo_diario_arriendo_vehiculo','arriendo_diario_vehiculo',default=np.nan)
    km_fee=pick(control,'costo_cop_km',default=np.nan)
    # La tarifa se lee de parámetros existentes; no asumir 5000 si el Excel no lo registra.
    for p in ('FII','FI'):
        par=phase_sheet(book,p,'Parametros')
        if not par.empty and {'parametro','valor'}<=set(par.columns):
            d={str(k):v for k,v in zip(par.parametro,par.valor)}
            if not np.isfinite(num(km_fee)):
                km_fee=pick(d,'cost_per_km',default=np.nan)
            if not np.isfinite(num(raw_fee)):
                raw_fee=pick(d,'vehicle_rental_daily',default=np.nan)
            vehicle_capacity=num(pick(d,'vehicle_capacity',default=2500),2500)
            break
    else:
        vehicle_capacity=2500

    coords=sheet(book,'Coordenadas_Nodos')
    if {'nodo_id','lat','lon'} <= set(coords.columns):
        coords=coords[['nodo_id','lat','lon']].copy()
        coords['nodo_id']=coords['nodo_id'].astype(str).str.replace(r'\.0$','',regex=True)
        to_numeric(coords,['lat','lon'])
        coords=coords[coords.lat.between(-90,90)&coords.lon.between(-180,180)].drop_duplicates('nodo_id')
    else:
        coords=pd.DataFrame(columns=['nodo_id','lat','lon'])

    available=[p for p in ('FI','FII') if phase_data[p]['detailed']]
    selected_phase=st.sidebar.selectbox('Fase principal para rutas',('Fase II','Fase I'),
                                         index=0)
    active='FI' if selected_phase=='Fase I' else 'FII'
    if active not in available and available:
        st.sidebar.info(f'No hay rutas detalladas de {active} en este Excel; se muestran las de {available[0]}.')
        active=available[0]

    refi=summaries['FI']['costo_total']; refii=summaries['FII']['costo_total']
    saving=refi-refii if np.isfinite(refi) and np.isfinite(refii) else np.nan
    percent=100*saving/refi if np.isfinite(saving) and refi>0 else np.nan
    availability_label=('FI y FII con rutas' if len(available)==2
                        else ('Información detallada disponible: '+', '.join(available)
                              if available else 'Solo resúmenes'))
    st.caption(f'Ejecución seleccionada: {source_name} · {availability_label}')

    names=['Resumen ejecutivo','Comparación de fases','Análisis económico','Mapa general',
           'Rutas individuales','Bodegas y asignación','Vehículos y rutas','Rendimiento',
           'Top cinco','Validación','Ahorros por movimiento',
           'Evolución y ahorros','Descargas e informe','Comparar instancias']
    tabs=st.tabs(names)
    with tabs[0]:
        a1,a2,a3,a4=st.columns(4)
        with a1: money_metric('Costo Fase I · COP/día',refi)
        with a2: money_metric('Costo Fase II · COP/día',refii)
        with a3: st.metric('Ahorro diario',fmt_money(saving))
        with a4: st.metric('Reducción',f'{fmt_num(percent,2)} %' if np.isfinite(percent) else 'No disponible')
        first_feasible=num(pick(control,'primera_factible_costo_fase1'))
        build_save=(first_feasible-refi if np.isfinite(first_feasible) and np.isfinite(refi) else np.nan)
        total_save=(first_feasible-refii if np.isfinite(first_feasible) and np.isfinite(refii) else np.nan)
        build_pct=(100*build_save/first_feasible if np.isfinite(build_save) and first_feasible>0 else np.nan)
        total_pct=(100*total_save/first_feasible if np.isfinite(total_save) and first_feasible>0 else np.nan)
        st.markdown('#### Ahorro de toda la ejecución · sin doble contabilización')
        extra=st.columns(4)
        for c,label,val in zip(extra,
             ['Primera solución factible','Ahorro constructivo Fase I',
              'Ahorro adicional Fase II','Ahorro total desde el primer factible'],
             [fmt_money(first_feasible),fmt_money(build_save),fmt_money(saving),fmt_money(total_save)]):
            with c: st.metric(label,val)
        if not np.isfinite(first_feasible):
            st.caption('Esta ejecución histórica no registra el primer costo factible. No se infiere el ahorro constructivo.')
        current=summaries['FII']
        mm=st.columns(3) + st.columns(3)
        for col,label,value,dec in zip(mm,['Bodegas abiertas','Vehículos','Rutas','Clientes','Operaciones','Kilómetros'],
                                       [current['bodegas'],current['vehiculos'],current['rutas'],current['clientes'],
                                        current['operaciones'],current['km']],[0,0,0,0,0,2]):
            with col: st.metric(label,fmt_num(value,dec))
        n1,n2,n3=st.columns(3)
        with n1: st.metric('Tiempo transporte (sumado)',
                            f"{fmt_num(current['minutos']/60,2)} h" if np.isfinite(current['minutos']) else 'No disponible')
        with n2: st.metric('Tiempo total de ejecución',
                            f"{fmt_num(pick(control,'tiempo_total_segundos'),2)} s")
        with n3: st.metric('Ahorro Fase II (%)',f'{fmt_num(percent,2)} %')
        fig=bar_compare(summaries,[('costo_total','Costo total')], 'Costo total diario por fase','COP/día')
        if fig: show_chart(fig,use_container_width=True)
        if np.isfinite(saving):
            if saving>1e-7:
                st.info(f'Lectura del resultado: la Fase II redujo el costo diario en '
                    f'{fmt_money(saving)} COP ({fmt_num(percent,2)} %) frente a la Fase I. '
                    f'La comparación corresponde a las dos soluciones oficiales, sin sumar mejoras de alternativas distintas.')
            elif abs(saving)<=1e-7:
                st.info('Lectura del resultado: la Fase II conservó el costo de la Fase I; '
                        'no se registró un ahorro adicional en la solución oficial.')
            else:
                st.warning('La solución final registra un costo superior al inicial; revisa la comparación exportada.')
        if len(available)<2:
            st.info('Algunos detalles de rutas solo están disponibles para la fase registrada.')

    with tabs[1]:
        rows=[]
        for field,label,unit in [('costo_total','Costo total','COP/día'),('bodegas_cost','Bodegas','COP/día'),
                         ('arriendo','Arriendo','COP/día'),('transporte','Transporte','COP/día'),
                         ('km','Distancia','km'),('rutas','Rutas','rutas'),('vehiculos','Vehículos','unidades'),
                         ('minutos','Tiempo de transporte','min')]:
            i=summaries['FI'][field]; f=summaries['FII'][field]
            rows.append({'Indicador':label,'Unidad':unit,'Fase I':i,'Fase II':f,
                         'Diferencia (II − I)':f-i if np.isfinite(i) and np.isfinite(f) else np.nan})
        cmp=pd.DataFrame(rows)
        data_table(cmp,use_container_width=True,hide_index=True)
        download_csv(cmp,'LogiData_comparacion_fases.csv','cmp_csv')
        download_table_xlsx(cmp,'LogiData_comparacion_fases.xlsx','cmp_xlsx')
        for defs,title,unit in [([('km','Kilómetros'),('rutas','Rutas')],'Recorridos y flota','km / rutas'),
                                ([('bodegas_cost','Bodegas'),('arriendo','Arriendo'),('transporte','Transporte')],
                                  'Comparación de componentes','COP/día')]:
            fig=bar_compare(summaries,defs,title,unit)
            if fig: show_chart(fig,use_container_width=True)
        st.caption('Análisis por separado: '+savings_commentary(summaries['FI']['km'],
            summaries['FII']['km'],'La distancia recorrida',unit='km')+' '+
            savings_commentary(summaries['FI']['rutas'],summaries['FII']['rutas'],
                               'El número de vehículos/rutas',unit='unidades'))

    with tabs[2]:
        st.subheader('Composición de la función objetivo')
        eco=pd.DataFrame([{'Fase':p.replace('FI','Fase I').replace('Fase II I','Fase II'),
                           'Bodegas':summaries[p]['bodegas_cost'],
                           'Arriendo de vehículos':summaries[p]['arriendo'],
                           'Transporte':summaries[p]['transporte']} for p in ('FI','FII')])
        eco['Fase']=['Fase I','Fase II']
        data_table(eco,use_container_width=True,hide_index=True)
        melted=eco.melt(id_vars='Fase',var_name='Concepto',value_name='COP diarios').dropna()
        if not melted.empty:
            fig=px.bar(melted,x='Fase',y='COP diarios',color='Concepto',barmode='stack',
                       color_discrete_map={'Bodegas':NAVY,'Arriendo de vehículos':ORANGE,'Transporte':TEAL},
                       title='Estructura de costos diarios')
            fig.update_layout(height=390)
            show_chart(fig,use_container_width=True)

        # Distribution of the real final cost (doughnut chart, same original data).
        selected_costs = eco.loc[eco['Fase']=='Fase II', ['Bodegas','Arriendo de vehículos','Transporte']]
        if not selected_costs.empty:
            costs = selected_costs.iloc[0].dropna()
            costs = costs[costs >= 0]
            if costs.sum() > 0:
                pie = go.Figure(go.Pie(labels=list(costs.index), values=list(costs.values),
                    hole=.64, sort=False, textinfo='percent', textposition='outside',
                    marker=dict(colors=[NAVY,VIOLET,TEAL][:len(costs)])))
                pie.update_layout(title='Participación de los costos · Fase II',height=370,
                    annotations=[dict(text='Costo<br>Fase II',x=.5,y=.5,showarrow=False,
                                      font=dict(size=15,color=NAVY))],
                    showlegend=True,legend=dict(orientation='h',y=-.07,x=.12))
                show_chart(pie,use_container_width=True)

        if np.isfinite(num(raw_fee)):
            st.caption(f'Arriendo unitario reportado: {fmt_money(raw_fee)} COP por ruta y día')
        else:
            st.warning('El Excel no indica tarifa unitaria de arriendo: no se atribuye una distinta por defecto.')
        if np.isfinite(num(km_fee)):
            st.caption(f'Tarifa variable reportada: {fmt_money(km_fee)} COP/km')
        for p in available:
            audit=verify_phase(phase_data[p],num(km_fee),vehicle_capacity)
            rc=audit[audit['Restricción / control']=='Reconciliación de función objetivo diaria']
            if not rc.empty and rc.iloc[0]['Resultado']=='Incumple':
                st.error(f'Los componentes económicos de {p} no cuadran. Revisa la hoja de validación.')
        if not eco[['Bodegas','Arriendo de vehículos','Transporte']].isna().all().all():
            delta=eco.set_index('Fase').loc['Fase II',['Bodegas','Arriendo de vehículos','Transporte']]-eco.set_index('Fase').loc['Fase I',['Bodegas','Arriendo de vehículos','Transporte']]
            differ=pd.DataFrame({'Componente':delta.index,'Diferencia COP/día':delta.to_numpy()})
            differ['Cambio (%)']=[(100*d/float(eco.set_index('Fase').loc['Fase I',c])
               if float(eco.set_index('Fase').loc['Fase I',c])!=0 else np.nan)
               for d,c in zip(differ['Diferencia COP/día'],differ['Componente'])]
            st.markdown('#### Cambios económicos al pasar de la Fase I a la Fase II')
            data_table(differ,use_container_width=True,hide_index=True)
            fig_delta=px.bar(differ,x='Componente',y='Diferencia COP/día',
                color='Componente',color_discrete_sequence=[NAVY,ORANGE,TEAL],
                text_auto='.2f',title='Cambio de cada componente (valores negativos = ahorros)')
            show_chart(fig_delta,use_container_width=True)
            changes = []
            for component,key in [('Transporte','transporte'),('Arriendo de vehículos','arriendo'),
                                  ('Bodegas','bodegas_cost')]:
                changes.append(savings_commentary(summaries['FI'][key],summaries['FII'][key],
                                                   'El costo de '+component.lower()))
            st.info('Análisis de la distribución: ' + ' '.join(changes) +
                ' Una subida en un componente puede compensarse con reducciones en otros; '
                'el ahorro total se calcula sobre el costo diario completo.')

        download_csv(eco,'LogiData_costos_diarios.csv','eco_csv')
        download_table_xlsx(eco,'LogiData_costos_diarios.xlsx','eco_xlsx')

    with tabs[3]:
        st.subheader(f'Mapa general · {active}')
        if active in available:
            r=phase_data[active]['routes']
            dp=sorted(r.bodega_id.unique().tolist())
            selected=st.multiselect('Filtrar bodegas abiertas',dp,default=dp,key='general_depots')
            rr=r[r.bodega_id.isin(selected)]
            st.caption('Selecciona un vehículo para ver sus pasos numerados sobre el mapa. '
                       'Las demás rutas siguen visibles para comparar los recorridos.')
            route_ids_map=rr.ruta_id.astype(str).tolist()
            highlighted=st.selectbox('Numerar las visitas del vehículo',
                route_ids_map if route_ids_map else ['Sin rutas'],
                format_func=lambda x: friendly_vehicle(x) if x!='Sin rutas' else 'Sin rutas',
                key='general_route_highlight')
            modo_offline=st.checkbox('Ver coordenadas si el mapa de calles no carga',
                value=False,key='mapa_general_coordenadas',help='Modo sin conexión, con ejes de latitud y longitud')
            fig=map_figure(rr,phase_data[active]['visits'],coords,
                'Rutas y bodegas abiertas',modo_coordenadas=modo_offline,
                map_order_route=highlighted if highlighted!='Sin rutas' else None)
            if fig:
                show_chart(fig,use_container_width=True)
                st.caption('Las conexiones son segmentos rectos esquemáticos entre coordenadas reales; no representan la geometría de las calles.')

            else: st.warning('No hay coordenadas suficientes en esta ejecución para visualizar las rutas.')
            if highlighted!='Sin rutas':
                steps=phase_data[active]['visits']
                steps=steps[steps.ruta_id.astype(str)==str(highlighted)].sort_values('orden')
                route_row=rr[rr.ruta_id.astype(str)==str(highlighted)]
                if not steps.empty and not route_row.empty:
                    st.markdown(f'#### Orden real de visita · {friendly_vehicle(highlighted)}')
                    st.caption('Salida y regreso a la bodega incluidos. Si un cliente tiene entrega y recogida, sus dos números se muestran juntos en el mapa.')
                    st.write('  →  '.join([('Bodega '+str(x.nodo_id) if x.operacion.startswith('DEPOT')
                          else str(int(x.orden))+'. Cliente '+str(x.nodo_id))
                          for _,x in steps.iterrows()]))
                    ordered_map=ordered_visit_table(steps,route_row.iloc[0].bodega_id)
                    data_table(ordered_map,use_container_width=True)
                    download_csv(ordered_map,'LogiData_orden_visitas_mapa.csv','general_order_csv')
                    download_table_xlsx(ordered_map,'LogiData_orden_visitas_mapa.xlsx','general_order_excel')
        else: st.info('El Excel no contiene rutas detalladas de la fase seleccionada.')

    with tabs[4]:
        if active in available:
            pdata=phase_data[active]
            rr=pdata['routes']
            depot_ids=sorted(rr.bodega_id.unique())
            chosen_depot=st.selectbox('Bodega', ['Todas']+depot_ids, key='route_depot')
            visible=rr if chosen_depot=='Todas' else rr[rr.bodega_id==chosen_depot]
            route_id=st.selectbox('Ruta / vehículo', visible.ruta_id.tolist(), key='route_detail',
                format_func=friendly_vehicle)
            if route_id is None:
                st.info('Selecciona al menos una bodega para mostrar una ruta.')
                st.info('No existen rutas para esa selección.')
                return
            r=rr[rr.ruta_id==route_id]
            visits=pdata['visits'][pdata['visits'].ruta_id==route_id].sort_values('orden')
            details=route_table(pdata)
            dr=details[details.ruta==route_id].iloc[0]
            cols=st.columns(5)
            for col,label,val in zip(cols,['Bodega','Visitas','Entregas / recogidas','Kilómetros','Costo diario ruta'],
                                     [dr.bodega,str(dr.visitas),f'{dr.unidades_entregadas:,.0f} / {dr.unidades_recogidas:,.0f}',
                                      fmt_num(dr.km,2),fmt_money(dr.costo_ruta_diario)]):
                with col: st.metric(label,str(val))
            modo_offline_ruta=st.checkbox('Ver coordenadas si el mapa de calles no carga',
                value=False,key='mapa_individual_coordenadas')
            fig=map_figure(r,pdata['visits'],coords,
                f'{friendly_route(route_id)} · salida y regreso a bodega {dr.bodega}',
                modo_coordenadas=modo_offline_ruta)
            if fig:
                show_chart(fig,use_container_width=True)

            else: st.info('Mapa no disponible por falta de coordenadas de las visitas. La tabla de secuencia sí es válida.')
            st.caption('Conexiones en línea recta: esquema del orden registrado, no geometría vial.')
            st.markdown('#### Recorrido completo, paso a paso')
            st.caption('Los números sobre el mapa muestran los pasos reales de atención. '
                'Una etiqueta como 2/5 indica dos visitas al mismo cliente; la tabla presenta ambas por separado.')
            full_route=readable_sequence(visits, max_chars=3000)
            st.markdown('**Secuencia:** ' + full_route)
            ordered_route=ordered_visit_table(visits,dr.bodega)
            data_table(ordered_route,use_container_width=True,hide_index=True)
            st.markdown('#### Carga transportada en cada visita')
            fig2=px.line(visits,x='orden',y='carga_despues',markers=True,title='Evolución de la carga por visita',
                         labels={'orden':'Orden de visita','carga_despues':'Unidades en vehículo'})
            fig2.add_hline(y=vehicle_capacity,line_dash='dash',line_color=RED,
                           annotation_text=f'Capacidad {vehicle_capacity:,.0f}')
            show_chart(fig2,use_container_width=True)
            download_csv(ordered_route,f'LogiData_orden_visitas_{active}_{route_id}.csv','route_csv')
            download_table_xlsx(ordered_route,f'LogiData_orden_visitas_{active}_{route_id}.xlsx','route_xlsx')
            if len(available)==2:
                other='FI' if active=='FII' else 'FII'
                st.caption('Puedes alternar las fases en el selector lateral. Cada recorrido se identifica dentro de su propia fase.')
        else: st.info('No hay rutas detalladas disponibles.')

    with tabs[5]:
        if active in available:
            p=phase_data[active]
            b=p['depots'].copy(); a=p['assignments'].copy()
            if not b.empty:
                b['utilizacion_pct']=(100*b.entrega_asignada/b.capacidad).replace([np.inf,-np.inf],np.nan)
                b['capacidad_disponible']=b.capacidad-b.entrega_asignada
                chosen=st.multiselect('Bodegas a comparar',sorted(b.bodega_id.unique()),
                                      default=sorted(b.bodega_id.unique()),key='depots_filter')
                b=b[b.bodega_id.isin(chosen)]
                data_table(b,use_container_width=True,hide_index=True)
                fig=px.bar(b,x='bodega_id',y=['entrega_asignada','capacidad_disponible'],
                           title='Capacidad de bodega utilizada y disponible',barmode='stack',
                           labels={'value':'Unidades','bodega_id':'Bodega','variable':'Categoría'},
                           color_discrete_sequence=[TEAL,'#CBD7E0'])
                show_chart(fig,use_container_width=True)
                download_csv(b,'LogiData_bodegas.csv','dep_csv')
                download_table_xlsx(b,'LogiData_bodegas.xlsx','dep_xlsx')
            if not a.empty:
                a['tipo_cliente']=np.select([(a.entrega>0)&(a.recogida>0),a.entrega>0,a.recogida>0],
                                    ['Mixto','Solo entrega','Solo recogida'],default='Otro')
                hist=a.groupby(['bodega_id','tipo_cliente'],as_index=False).size()
                fig=px.bar(hist,x='bodega_id',y='size',color='tipo_cliente',barmode='stack',
                          color_discrete_sequence=[VIOLET,BLUE,ORANGE],
                          title='Clientes atendidos por bodega y tipo',
                          labels={'size':'Clientes','bodega_id':'Bodega','tipo_cliente':'Tipo'})
                show_chart(fig,use_container_width=True)
                data_table(a,use_container_width=True,hide_index=True)
                download_csv(a,'LogiData_asignaciones.csv','asig_csv')
                download_table_xlsx(a,'LogiData_asignaciones.xlsx','asig_xlsx')
        else: st.info('Sin asignaciones o bodegas detalladas de la fase seleccionada.')

        # Comparación estructural de las bodegas y reasignaciones entre soluciones OFICIALES.
        if all(phase_data[x]['detailed'] for x in ('FI','FII')):
            f1=phase_data['FI']['assignments']
            f2=phase_data['FII']['assignments']
            d1=phase_data['FI']['depots']
            d2=phase_data['FII']['depots']
            if not f1.empty and not f2.empty:
                st.markdown('#### Cambios de localización y asignación entre fases')
                old=set(d1.bodega_id) if not d1.empty else set(phase_data['FI']['routes'].bodega_id)
                new=set(d2.bodega_id) if not d2.empty else set(phase_data['FII']['routes'].bodega_id)
                s1=f1[['cliente_id','bodega_id']].rename(columns={'bodega_id':'bodega_fase_I'})
                s2=f2[['cliente_id','bodega_id']].rename(columns={'bodega_id':'bodega_fase_II'})
                diff=s1.merge(s2,on='cliente_id',how='outer',validate='one_to_one')
                diff=diff[diff.bodega_fase_I!=diff.bodega_fase_II]
                c_a,c_b,c_c=st.columns(3)
                with c_a: st.metric('Bodegas nuevas',len(new-old))
                with c_b: st.metric('Bodegas cerradas',len(old-new))
                with c_c: st.metric('Clientes reasignados',len(diff))
                if new-old: st.info('Aperturas identificadas: '+', '.join(sorted(new-old)))
                if old-new: st.info('Cierres identificados: '+', '.join(sorted(old-new)))
                if not diff.empty:
                    data_table(diff,use_container_width=True,hide_index=True)
                    download_csv(diff,'LogiData_clientes_trasladados_entre_fases.csv','transferred_csv')
                elif old==new: st.caption('No se detectaron cambios de bodegas ni de asignaciones entre las soluciones oficiales.')
            if not d1.empty and not d2.empty:
                oldcols=d1[['bodega_id','capacidad','entrega_asignada']].rename(columns={'entrega_asignada':'entrega_FI'})
                newcols=d2[['bodega_id','capacidad','entrega_asignada']].rename(columns={'entrega_asignada':'entrega_FII'})
                changes=oldcols.merge(newcols[['bodega_id','entrega_FII']],on='bodega_id',how='outer').fillna({'entrega_FI':0,'entrega_FII':0})
                c_fig=px.bar(changes,x='bodega_id',y=['entrega_FI','entrega_FII'],barmode='group',
                             title='Demanda de entregas asignada por bodega · FI vs FII',
                             labels={'value':'Unidades','bodega_id':'Bodega','variable':'Fase'},
                             color_discrete_sequence=[BLUE,TEAL])
                show_chart(c_fig,use_container_width=True)

    with tabs[6]:
        routes_all=pd.concat([route_table(phase_data[p],vehicle_capacity) for p in available],ignore_index=True) if available else pd.DataFrame()
        if not routes_all.empty:
            phases_sel=st.multiselect('Fases',sorted(routes_all.fase.unique()),
                                    default=sorted(routes_all.fase.unique()),key='fleet_phases')
            depots_sel=st.multiselect('Bodegas',sorted(routes_all.bodega.unique()),
                                    default=sorted(routes_all.bodega.unique()),key='fleet_dep')
            eligible=routes_all[routes_all.fase.isin(phases_sel)&routes_all.bodega.isin(depots_sel)]
            route_select=st.multiselect('Vehículos / rutas',sorted(eligible.vehiculo.unique()),
                                      default=sorted(eligible.vehiculo.unique()),key='fleet_vehicle_filter',
                                      format_func=friendly_vehicle)
            visible=eligible[eligible.vehiculo.isin(route_select)]
            data_table(visible,use_container_width=True,hide_index=True)
            download_csv(visible,'LogiData_rutas_vehiculos.csv','fleet_csv')
            download_table_xlsx(visible,'LogiData_rutas_vehiculos.xlsx','fleet_xlsx')
            counts=visible.groupby(['fase','bodega'],as_index=False).agg(vehiculos=('ruta','size'))
            if not counts.empty:
                fig=px.bar(counts,x='bodega',y='vehiculos',color='fase',barmode='group',
                           title='Vehículos por instalación',color_discrete_map={'FI':BLUE,'FII':TEAL})
                show_chart(fig,use_container_width=True)
            if not visible.empty and {'vehiculo','km','costo_ruta_diario'} <= set(visible.columns):
                # Nombres breves también en el hover; sin códigos internos de ruta.
                visible_plot=visible.copy()
                visible_plot['vehiculo']=visible_plot.vehiculo.map(friendly_vehicle)
                distance_plot = px.bar(visible_plot, y='vehiculo', x='km', color='fase',
                    orientation='h', barmode='group', color_discrete_map={'FI':BLUE,'FII':VIOLET},
                    labels={'vehiculo':'Vehículo','km':'Kilómetros','fase':'Fase'},
                    title='Distancia recorrida por vehículo')
                distance_plot.update_layout(height=max(345, 30*visible.vehiculo.nunique()+100))
                show_chart(distance_plot,use_container_width=True)
                st.markdown('#### ¿Qué rutas explican los costos y las distancias?')
                details_cost=visible_plot.dropna(subset=['km','costo_ruta_diario']).copy()
                if not details_cost.empty:
                    route_scatter=px.scatter(details_cost,x='km',y='costo_ruta_diario',
                        color='bodega',symbol='fase',hover_name='vehiculo',
                        hover_data=['visitas','minutos','arriendo','costo_transporte'],
                        labels={'km':'Distancia recorrida (km)',
                                'costo_ruta_diario':'Costo diario de ruta (COP)',
                                'bodega':'Bodega','fase':'Fase'},
                        title='Distancia frente a costo diario por vehículo')
                    route_scatter.update_traces(marker=dict(size=12,line=dict(width=1,color='#FFFFFF')))
                    route_scatter.update_layout(height=390)
                    show_chart(route_scatter,use_container_width=True)
                    costs_bar=px.bar(details_cost.sort_values('costo_ruta_diario'),
                        x='costo_ruta_diario',y='vehiculo',color='fase',orientation='h',
                        hover_data=['bodega','km','minutos'],
                        labels={'costo_ruta_diario':'Costo diario de ruta (COP)','vehiculo':'Vehículo'},
                        title='Costo diario por ruta · transporte y arriendo incluidos',
                        color_discrete_map={'FI':BLUE,'FII':TEAL})
                    costs_bar.update_layout(height=max(335,34*len(details_cost)+100))
                    show_chart(costs_bar,use_container_width=True)
                    st.markdown('#### Composición y actividad de cada vehículo')
                    # Un vehículo aporta un arriendo diario más el costo variable
                    # de sus arcos. El costo de bodega queda fuera de esta gráfica.
                    has_cols={'vehiculo','costo_transporte','arriendo','fase'} <= set(details_cost.columns)
                    if has_cols:
                        split=details_cost[['vehiculo','fase','costo_transporte','arriendo']].melt(
                            id_vars=['vehiculo','fase'],value_vars=['costo_transporte','arriendo'],
                            var_name='Concepto',value_name='COP/día')
                        split['Concepto']=split['Concepto'].replace({
                            'costo_transporte':'Transporte','arriendo':'Arriendo'})
                        stacked=px.bar(split,x='vehiculo',y='COP/día',color='Concepto',
                            facet_row='fase' if len(details_cost.fase.unique())>1 else None,
                            color_discrete_map={'Transporte':TEAL,'Arriendo':ORANGE},
                            title='Costo diario de cada vehículo: transporte + arriendo')
                        stacked.update_layout(height=420)
                        show_chart(stacked,use_container_width=True)
                    if {'unidades_entregadas','unidades_recogidas'} <= set(details_cost.columns):
                        operations=details_cost[['vehiculo','fase','unidades_entregadas','unidades_recogidas']].melt(
                            id_vars=['vehiculo','fase'],
                            value_vars=['unidades_entregadas','unidades_recogidas'],
                            var_name='Operación',value_name='Unidades')
                        operations['Operación']=operations['Operación'].replace({
                            'unidades_entregadas':'Entregas','unidades_recogidas':'Recogidas'})
                        bar_ops=px.bar(operations,y='vehiculo',x='Unidades',color='Operación',
                            orientation='h',barmode='group',
                            color_discrete_map={'Entregas':BLUE,'Recogidas':ORANGE},
                            title='Volumen de entregas y recogidas por vehículo')
                        bar_ops.update_layout(height=max(320,35*details_cost.vehiculo.nunique()+95))
                        show_chart(bar_ops,use_container_width=True)
                    largest=details_cost.loc[details_cost.costo_ruta_diario.idxmax()]
                    st.caption(f'Lectura de las rutas filtradas: {largest.vehiculo}, de la bodega '
                        f'{largest.bodega}, presenta el mayor costo diario: '
                        f'{fmt_money(largest.costo_ruta_diario)} COP para '
                        f'{fmt_num(largest.km,2)} km. En esta cifra se incluye transporte '
                        'y arriendo; el costo de bodega se analiza por separado.')
            st.caption('Vehículo lógico: una ruta diaria por unidad; la salida y el retorno NO se cuentan como dos arriendos.')
            st.markdown('#### Consulta de rutas y operaciones por vehículo')
            f_sel=st.selectbox('Fase para detalles',available,index=available.index(active),key='fleet_phase_detail')
            phase_r=phase_data[f_sel]['routes']
            phase_v=phase_data[f_sel]['visits']
            choices=sorted(phase_r.bodega_id.unique().tolist())
            dep_selected=st.selectbox('Bodega de origen', ['Todas']+choices,key='fleet_depot_detail')
            filtered_routes=phase_r if dep_selected=='Todas' else phase_r[phase_r.bodega_id==dep_selected]
            ids=sorted(filtered_routes.ruta_id.unique().tolist())
            route_selected=st.selectbox('Vehículo / ruta', ['Todos']+ids,key='fleet_route_detail',
                format_func=lambda x: x if x=='Todos' else friendly_vehicle(x))
            r_ids=ids if route_selected=='Todos' else [route_selected]
            rows=phase_v[phase_v.ruta_id.isin(r_ids)].copy().sort_values(['ruta_id','orden'])
            if not rows.empty:
                rows.insert(0,'fase',f_sel)
                rows.insert(1,'vehiculo',rows.ruta_id.map(lambda x:f'{f_sel}-{x}'))
                data_table(rows)
                download_csv(rows,f'LogiData_{f_sel}_detalle_paradas.csv','fleet_stops_csv')
                download_table_xlsx(rows,f'LogiData_{f_sel}_detalle_paradas.xlsx','fleet_stops_xlsx')
            st.markdown('#### Clústeres, segmentos y emparejamientos')
            kind=st.selectbox('Tabla de construcción',['Clústeres','Miembros de clúster',
              'Segmentos','Visitas de segmento','Emparejamientos'],key='structure_select')
            keymap={'Clústeres':'clusters','Miembros de clúster':'members',
              'Segmentos':'segments','Visitas de segmento':'segment_visits','Emparejamientos':'pairings'}
            structure=phase_data[f_sel][keymap[kind]].copy()
            if 'bodega_id' in structure and dep_selected!='Todas':
                structure=structure[structure.bodega_id.astype(str)==str(dep_selected)]
            if structure.empty: st.info('No existen registros para la selección.')
            else:
                data_table(structure)
                download_csv(structure,f'LogiData_{f_sel}_{keymap[kind]}.csv','structure_csv')
                download_table_xlsx(structure,f'LogiData_{f_sel}_{keymap[kind]}.xlsx','structure_xlsx')

        else: st.info('El Excel no tiene rutas detalladas.')

    with tabs[7]:
        st.subheader('Rendimiento y trazabilidad')
        f1=num(pick(control,'tiempo_fase1_segundos',default=np.nan))
        f2=num(pick(control,'tiempo_fase2_segundos',default=np.nan))
        total=num(pick(control,'tiempo_total_segundos',default=np.nan))
        cl=st.columns(3) + st.columns(2)
        vals=[('Tiempo total',total,'s'),('Fase I',f1,'s'),('Fase II',f2,'s'),
              ('Construcciones',num(pick(control,'construcciones_fase1',default=np.nan)),'intentos'),
              ('Mejoras Fase II',num(pick(control,'mejoras_fase2',default=np.nan)),'movimientos')]
        for c,(label,value,u) in zip(cl,vals):
            with c: st.metric(label,(f'{fmt_num(value,2)} s' if u=='s' else fmt_num(value,0))
                if np.isfinite(value) else 'No registrado')
        for name,label in [('FaseI_Lotes','Lotes y evolución de Fase I'),
                           ('Tiempos_Movimientos','Tiempos y evaluaciones por vecindario'),
                           ('Movimientos_FaseII','Movimientos aceptados y evolución del costo'),
                           ('FaseI_Evolucion','Cada mejora del mejor costo global de Fase I')]:
            df=sheet(book,name)
            if not df.empty and 'candidato' in df.columns:
                opt=sorted(df['candidato'].dropna().astype(str).unique().tolist())
                keep=st.multiselect(f'Alternativas · {label}',opt,default=opt,
                    key=f'filter_candidate_{name}',format_func=friendly_alternative)
                df=df[df['candidato'].astype(str).isin(keep)]
            sweep_column=next((x for x in ('barrido','ciclo') if x in df.columns),None)
            if not df.empty and sweep_column:
                allowed=sorted(pd.to_numeric(df[sweep_column],errors='coerce').dropna().unique().tolist())
                if len(allowed)>1:
                    selected=st.multiselect(f'Barridos/ciclos · {name}',allowed,
                                            default=allowed,key=f'filter_sweep_{name}')
                    df=df[df[sweep_column].isin(selected)]
            st.markdown('#### '+label)
            if df.empty:
                st.info('Esta ejecución no contiene registros de esa medición.')
                continue
            data_table(df,use_container_width=True,hide_index=True)
            download_csv(df,f'LogiData_{name}.csv',f'dl_{name}')
            if name=='FaseI_Lotes' and {'lote','mejor_costo'}<=set(df.columns):
                fig=px.line(df,x='lote',y='mejor_costo',markers=True,
                            title='Evolución del mejor costo global en Fase I')
                show_chart(fig,use_container_width=True)
            if name=='Tiempos_Movimientos' and {'movimiento','segundos'}<=set(df.columns):
                totals=df.groupby('movimiento',as_index=False).agg(
                    segundos=('segundos','sum'),evaluaciones=('evaluaciones','sum'))
                fig=px.bar(totals,x='movimiento',y='segundos',text_auto='.2f',
                           title='Tiempo acumulado por movimiento (cinco candidatas)')
                show_chart(fig,use_container_width=True)
            if name=='Movimientos_FaseII' and {'orden_global','costo_despues'}<=set(df.columns):
                display_trace=df.copy()
                display_trace['candidato']=display_trace['candidato'].map(friendly_alternative)
                fig=px.line(display_trace,x='orden_global',y='costo_despues',color='candidato',markers=True,
                            title='Evolución del costo después de mejoras aceptadas')
                show_chart(fig,use_container_width=True)
                if {'movimiento','ahorro'}<=set(df.columns):
                    savings=df.groupby('movimiento',as_index=False).agg(
                        mejoras=('ahorro','size'),ahorro_cop=('ahorro','sum'))
                    fig_saved=px.bar(savings,x='movimiento',y='ahorro_cop',text='mejoras',
                        title='Ahorro total por movimiento · etiqueta = mejoras aceptadas')
                    show_chart(fig_saved,use_container_width=True)
                    st.caption('Interpretación: las barras representan ahorros de movimientos registrados '
                        'en las candidatas filtradas, que pueden haberse optimizado de forma independiente. '
                        'No deben sumarse directamente para afirmar el ahorro oficial entre fases.')
        st.markdown('#### Resumen de ejecución')
        claves=['tiempo_total_segundos','tiempo_fase1_segundos','tiempo_fase2_segundos',
                'construcciones_fase1','construcciones_factibles_fase1','construcciones_fallidas_fase1',
                'lotes_ejecutados','mejoras_fase2','evaluaciones_fase2','evaluaciones_apertura',
                'ciclos_adicionales_realizados']
        resumen_control=pd.DataFrame([{'Indicador':k.replace('_',' ').capitalize(),'Valor':control[k]}
                  for k in claves if k in control and pd.notna(control[k])])
        data_table(resumen_control)

    with tabs[8]:
        st.subheader('Las mejores construcciones y sus refinamientos')
        top=sheet(book,'Top5_FaseI')
        cand=sheet(book,'Evaluacion_Candidatos')
        if cand.empty: cand=sheet(book,'Candidatas_FaseII')
        if not top.empty:
            st.markdown('**Top inicial de Fase I**')
            data_table(top,use_container_width=True,hide_index=True)
            download_csv(top,'LogiData_top5.csv','top_csv')
        else:
            st.info('El archivo histórico no conserva las cinco soluciones iniciales. No se reconstruyen a partir de la solución ganadora.')
        if not cand.empty:
            st.markdown('**Resultados independientes de Fase II**')
            data_table(cand,use_container_width=True,hide_index=True)
            download_csv(cand,'LogiData_candidatos_fase2.csv','cand_csv')
            start=next((c for c in ['costo_inicial','costo_inicio'] if c in cand),None)
            end=next((c for c in ['costo_final'] if c in cand),None)
            idcol=next((c for c in ['candidato','posicion'] if c in cand),None)
            if start and end and idcol:
                cand_plot=cand.copy()
                cand_plot[idcol]=cand_plot[idcol].map(friendly_alternative)
                fig=px.bar(cand_plot,x=idcol,y=[start,end],barmode='group',
                          title='Costo diario inicial y refinado por alternativa')
                show_chart(fig,use_container_width=True)
        else: st.info('No se exportó la comparación de candidatas.')
        winner=pick(comparison,'mejor_candidato_fase2',default=np.nan)
        if pd.notna(winner):
            st.markdown(f'Alternativa seleccionada: {friendly_alternative(winner)}')

    with tabs[9]:
        results=[]
        for p in ('FI','FII'):
            results.append(verify_phase(phase_data[p],num(km_fee),vehicle_capacity))
        table=pd.concat(results,ignore_index=True)
        data_table(table,use_container_width=True,hide_index=True)
        if 'Incumple' in table.Resultado.values:
            st.error('Se detectaron incumplimientos en los registros de la solución. Revísalos antes de citar los resultados.')
        elif 'No verificable' in table.Resultado.values:
            st.warning('Algunos controles no pueden comprobarse por falta de estructuras en el Excel histórico.')
        else: st.success('Todos los controles efectuados sobre las estructuras exportadas han pasado.')
        st.markdown('#### Salidas, regresos y alquiler por bodega')
        for p in available:
            b=phase_data[p]['depots']
            if not b.empty:
                st.write('**'+p+'**')
                columns=['bodega_id','vehiculos_utilizados','vehiculos_salen','vehiculos_regresan',
                         'costo_arriendo_diario','costo_diario_bodega']
                data_table(b[[x for x in columns if x in b.columns]],use_container_width=True,hide_index=True)
        download_csv(table,'LogiData_auditoria_restricciones.csv','validation_csv')
        download_table_xlsx(table,'LogiData_auditoria_restricciones.xlsx','validation_xlsx')

    with tabs[10]:
        st.subheader('Búsqueda local · movimientos y apertura de bodegas')
        st.caption('Se muestran únicamente evaluaciones, mejoras y tiempos registrados por el ejecutor. Los costos son COP por jornada.')
        mv=sheet(book,'Movimientos_FaseII')
        timing=sheet(book,'Tiempos_Movimientos')
        if mv.empty:
            st.info('Este archivo no conserva trazabilidad de movimientos aceptados. No es posible reconstruirla a partir del costo final.')
        else:
            mv=mv.copy()
            mv['nombre_movimiento']=mv['movimiento'].replace({
                'opening':'Apertura de bodegas','two_opt':'2-opt','relocate':'Relocate',
                'swap':'Swap','fusion':'Fusión','reassignment':'Reasignación'})
            kinds=sorted(mv.nombre_movimiento.dropna().unique())
            chosen=st.multiselect('Movimientos',kinds,default=kinds,key='opening_move_filter')
            labels=sorted(mv.candidato.dropna().astype(str).unique()) if 'candidato' in mv else []
            winner_rank=pick(comparison,'mejor_candidato_fase2',default=np.nan)
            winner_label=(str(int(winner_rank)) if np.isfinite(num(winner_rank)) else None)
            official=[x for x in labels if x==winner_label or x.startswith('ADICIONAL-')]
            select_labels=st.multiselect('Alternativas / ciclos',labels,
                                default=(official or labels),key='opening_candidate_filter',
                                format_func=friendly_alternative)
            if official and set(select_labels)!=set(official):
                st.caption('Los ahorros de candidatas independientes son exploratorios y NO se suman al ahorro oficial FI→FII.')
            visible=mv[mv.nombre_movimiento.isin(chosen)&mv.candidato.astype(str).isin(select_labels)]
            c1,c2,c3=st.columns(3)
            with c1: st.metric('Mejoras aceptadas',len(visible))
            with c2: st.metric('Ahorro de los movimientos filtrados (exploratorio)',fmt_money(visible.ahorro.sum()))
            with c3: st.metric('Aperturas aceptadas',int((visible.movimiento=='opening').sum()))
            data_table(visible,use_container_width=True,hide_index=True)
            download_csv(visible,'LogiData_mejoras_locales.csv','opening_moves_csv')
            download_table_xlsx(visible,'LogiData_mejoras_locales.xlsx','opening_moves_xlsx')
            if not visible.empty:
                change=visible.sort_values('orden_global')
                fig=px.scatter(change,x='orden_global',y='costo_despues',color='nombre_movimiento',
                               hover_data=['costo_antes','ahorro'],
                               title='Costo diario después de cada mejora aceptada',
                               labels={'orden_global':'Orden de la mejora','costo_despues':'COP/día'})
                show_chart(fig,use_container_width=True)

            opened=visible[visible.movimiento=='opening']
            st.markdown('#### Efecto económico de nuevas bodegas')
            if opened.empty:
                st.info('No se registraron aperturas aceptadas en los filtros actuales. Esto no significa que no se hayan evaluado alternativas.')
            else:
                cols=['candidato','pasada','detalle','bodegas_abiertas','clientes_trasladados',
                      'bodegas_costo_antes','bodegas_costo_despues',
                      'arriendo_antes','arriendo_despues','transporte_antes','transporte_despues',
                      'km_antes','km_despues','rutas_antes','rutas_despues',
                      'costo_antes','costo_despues','ahorro']
                data_table(opened[[x for x in cols if x in opened.columns]],use_container_width=True,hide_index=True)
                if {'bodegas_costo_antes','bodegas_costo_despues','arriendo_antes','arriendo_despues',
                        'transporte_antes','transporte_despues'}<=set(opened):
                    deltas=pd.DataFrame({'Componente':['Bodegas','Arriendo','Transporte'],
                        'Cambio COP/día':[
                            (opened.bodegas_costo_despues-opened.bodegas_costo_antes).sum(),
                            (opened.arriendo_despues-opened.arriendo_antes).sum(),
                            (opened.transporte_despues-opened.transporte_antes).sum()]})
                    fig=px.bar(deltas,x='Componente',y='Cambio COP/día',text_auto='.2f',
                               color='Componente',color_discrete_sequence=[NAVY,ORANGE,TEAL],
                               title='Incrementos y ahorros por componente tras aperturas aceptadas')
                    show_chart(fig,use_container_width=True)
                    st.caption('Valores negativos = ahorro. La suma de cambios corresponde a las aperturas visibles, no necesariamente al ahorro total de la Fase II.')
        st.markdown('#### Presupuestos, evaluaciones y tiempos por movimiento')
        if timing.empty:
            st.info('No hay tiempos por movimiento en este Excel.')
        else:
            data_table(timing,use_container_width=True,hide_index=True)
            download_csv(timing,'LogiData_tiempos_movimientos.csv','opening_timing_csv')
            for field,label in [('evaluaciones','Evaluaciones'),('segundos','Tiempo (s)'),('mejoras','Mejoras aceptadas')]:
                if field in timing:
                    data=timing.groupby('movimiento',as_index=False)[field].sum()
                    fig=px.bar(data,x='movimiento',y=field,color='movimiento',
                               title=f'{label} por movimiento',color_discrete_sequence=PALETTE)
                    show_chart(fig,use_container_width=True)
        if not timing.empty and 'movimiento' in timing:
            st.markdown('#### Balance de movimientos y ahorro observado')
            timing0=timing.copy()
            for metric in ('segundos','evaluaciones','mejoras'):
                if metric not in timing0: timing0[metric]=np.nan
            perf=timing0.groupby('movimiento',as_index=False).agg(
                tiempo_segundos=('segundos','sum'),evaluaciones=('evaluaciones','sum'),
                mejoras_registradas=('mejoras','sum'))
            if not mv.empty and {'movimiento','ahorro'}<=set(mv):
                accepted=mv.groupby('movimiento',as_index=False).agg(
                    mejoras_aceptadas=('ahorro','size'),ahorro_cop=('ahorro','sum'))
                perf=perf.merge(accepted,on='movimiento',how='left')
            else:
                perf['mejoras_aceptadas']=0;perf['ahorro_cop']=0.0
            perf['mejoras_aceptadas']=perf['mejoras_aceptadas'].fillna(0)
            perf['ahorro_cop']=perf['ahorro_cop'].fillna(0.0)
            perf['ahorro_por_segundo']=np.where(perf.tiempo_segundos>0,
                perf.ahorro_cop/perf.tiempo_segundos,np.nan)
            names={'opening':'Apertura','two_opt':'2-opt','relocate':'Relocate',
                'swap':'Swap','fusion':'Fusión','reassignment':'Reasignación'}
            perf['movimiento']=perf.movimiento.replace(names)
            data_table(perf)
            st.caption('Ahorros de todas las candidatas evaluadas: no sumarlos al ahorro oficial entre fases.')
            download_csv(perf,'LogiData_rendimiento_por_movimiento.csv','perf_mov_csv')
            download_table_xlsx(perf,'LogiData_rendimiento_por_movimiento.xlsx','perf_mov_xlsx')
            for field,label in [('ahorro_cop','Ahorro observado por movimiento'),
                                ('tiempo_segundos','Tiempo computacional por movimiento')]:
                f=px.bar(perf,y='movimiento',x=field,orientation='h',text_auto='.2f',
                    title=label,color='movimiento',color_discrete_sequence=PALETTE)
                show_chart(f,use_container_width=True)
        st.markdown('#### Ciclos adicionales')
        cycles=sheet(book,'Ciclos_Adicionales')
        if cycles.empty: st.caption('No se registraron ciclos adicionales en este archivo.')
        else:
            data_table(cycles,use_container_width=True,hide_index=True)
            download_csv(cycles,'LogiData_ciclos_adicionales.csv','opening_cycles_csv')


    with tabs[11]:
        st.subheader('Evolución real de la función objetivo')
        st.caption('La trayectoria constructiva registra nuevos mínimos globales de Fase I. '
                   'Los cinco candidatos de Fase II son optimizaciones independientes: '
                   'sus ahorros NO se suman para calcular el ahorro oficial.')
        stages=sheet(book,'Ahorros_Etapas')
        build_trace=sheet(book,'FaseI_Evolucion')
        move_trace=sheet(book,'Movimientos_FaseII')
        candidates=sheet(book,'Evaluacion_Candidatos')
        if not stages.empty:
            data_table(stages,use_container_width=True,hide_index=True)
        if not stages.empty and np.isfinite(first_feasible) and np.isfinite(refi) and np.isfinite(refii):
            cascada=go.Figure(go.Waterfall(
                x=['Primera factible','Mejora Fase I','Mejora Fase II','Costo final'],
                measure=['absolute','relative','relative','total'],
                y=[first_feasible,refi-first_feasible,refii-refi,0],
                connector={'line':{'color':'#A5B5C0','width':1}},
                decreasing={'marker':{'color':TEAL}},
                increasing={'marker':{'color':RED}},
                totals={'marker':{'color':NAVY}},
                text=[fmt_money(first_feasible),fmt_money(build_save),fmt_money(saving),fmt_money(refii)]))
            cascada.update_layout(title='Cascada de costos diarios · sin doble contabilización',
                yaxis_title='COP/día',height=405,showlegend=False)
            show_chart(cascada,use_container_width=True)
            st.info(f'Cómo interpretar la cascada: la Fase I produjo '
                f'{fmt_money(build_save)} COP/día de ahorro desde la primera solución factible. '
                f'La Fase II añadió {fmt_money(saving)} COP/día. '
                f'El ahorro total desde esa primera solución es {fmt_money(total_save)} COP/día '
                f'({fmt_num(total_pct,2)} %); estos importes no se suman dos veces.')

            download_csv(stages,'LogiData_ahorros_etapas.csv','stage_savings_export')
            if {'etapa','ahorro'}<=set(stages.columns):
                chart=px.bar(stages[stages.etapa!='Total de la ejecución'],x='etapa',y='ahorro',
                    text_auto='.2f',title='Ahorros no superpuestos · COP/día',
                    color='etapa',color_discrete_sequence=[TEAL,BLUE])
                show_chart(chart,use_container_width=True)

        else:
            st.info('El Excel histórico no conserva el primer factible; la mejora en Fase I no se reconstruye artificialmente.')
        if not build_trace.empty and {'mejor_costo','orden'}<=set(build_trace.columns):
            st.markdown('#### Evolución del mínimo global en Fase I')
            fig_build=px.line(build_trace,x='orden',y='mejor_costo',markers=True,
                hover_data=[x for x in ('lote','semilla_principal','cierre','segundos_desde_inicio')
                            if x in build_trace],
                title='Costos realmente mejorados durante las construcciones')
            fig_build.update_layout(yaxis_title='COP/día',xaxis_title='Número de mejora aceptada del mejor global')
            show_chart(fig_build,use_container_width=True)

            data_table(build_trace,use_container_width=True,hide_index=True)
        elif not sheet(book,'FaseI_Lotes').empty:
            batches=sheet(book,'FaseI_Lotes')
            if {'lote','mejor_costo'}<=set(batches):
                show_chart(px.line(batches,x='lote',y='mejor_costo',markers=True,
                               title='Mejor costo al cerrar cada lote'),use_container_width=True)
        if not move_trace.empty and {'candidato','orden_global','costo_despues'}<=set(move_trace):
            st.markdown('#### Evolución por candidata y ciclo adicional')
            labels=sorted(move_trace.candidato.astype(str).unique())
            chosen=st.selectbox('Seleccionar alternativa o ciclo',labels,key='trace_candidate_selector',
                format_func=friendly_alternative)
            sub=move_trace[move_trace.candidato.astype(str)==chosen].sort_values('orden_global')
            if not sub.empty:
                rows=[{'secuencia':0,'costo':num(sub.iloc[0].costo_antes),'movimiento':'Inicio'}]
                rows.extend({'secuencia':i,'costo':num(row.costo_despues),'movimiento':row.movimiento}
                            for i,row in enumerate(sub.itertuples(index=False),1))
                timeline=pd.DataFrame(rows)
                fig_local=px.area(timeline,x='secuencia',y='costo',markers=True,
                    hover_data=['movimiento'],title=f'Evolución de {friendly_alternative(chosen)} · mejoras aceptadas')
                fig_local.update_traces(line=dict(color=VIOLET,width=3),fillcolor='rgba(124,94,201,.12)')
                show_chart(fig_local,use_container_width=True)

                data_table(sub,use_container_width=True,hide_index=True)
        if not candidates.empty:
            st.markdown('#### Comparación de resultados independientes')
            data_table(candidates,use_container_width=True,hide_index=True)


    def crear_informe_pdf():
        """Resumen generado desde la ejecución cargada; sin optimizar ni crear datos."""
        try:
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.pagesizes import A4
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
        except ImportError:
            return None
        buffer=BytesIO()
        pdf=SimpleDocTemplate(buffer,pagesize=A4,rightMargin=42,leftMargin=42,topMargin=42,bottomMargin=42)
        styles=getSampleStyleSheet()
        content=[Paragraph('LOGIDATA | Resumen ejecutivo',styles['Title']),Spacer(1,13),
                 Paragraph(f'Ejecución: {source_name}',styles['Normal']),Spacer(1,15)]
        c1,c2=refi,refii
        table_rows=[['Indicador','Valor'],['Costo Fase I',fmt_money(c1)+' COP/día'],
                ['Costo Fase II',fmt_money(c2)+' COP/día'],
                ['Ahorro Fase II',fmt_money(saving)+' COP/día'],
                ['Reducción',fmt_num(percent,2)+' %'],
                ['Primera factible FI',fmt_money(first_feasible)+' COP/día'],
                ['Bodegas finales',fmt_num(summaries['FII']['bodegas'],0)],
                ['Vehículos finales',fmt_num(summaries['FII']['vehiculos'],0)],
                ['Kilómetros finales',fmt_num(summaries['FII']['km'],2)+' km'],
                ['Tiempo total',fmt_num(pick(control,'tiempo_total_segundos'),2)+' s']]
        t=Table(table_rows,colWidths=[235,255],repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor(NAVY)),
                               ('TEXTCOLOR',(0,0),(-1,0),colors.white),
                               ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F1F5F8')]),
                               ('BOTTOMPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),8)]))
        content+=[t,Spacer(1,16)]
        if np.isfinite(num(c1)) and np.isfinite(num(c2)):
            fig,ax=plt.subplots(figsize=(7,3))
            ax.bar(['Fase I','Fase II'],[c1,c2],color=[BLUE,TEAL])
            ax.set_ylabel('COP diarios');ax.set_title('Comparación del costo total diario')
            ax.ticklabel_format(axis='y',style='plain');fig.tight_layout()
            im=BytesIO();fig.savefig(im,format='png',dpi=135,bbox_inches='tight');plt.close(fig)
            im.seek(0);content.append(Image(im,width=480,height=205))
        content += [Spacer(1,14),Paragraph('Bodegas + arriendo vehicular + transporte = costo diario total.',styles['Normal']),
            Spacer(1,10),Paragraph('Mapas: trazado esquemático entre coordenadas, no geometría vial.',styles['Italic'])]
        pdf.build(content);return buffer.getvalue()


    with tabs[12]:
        st.subheader('Exportación de resultados')
        st.caption('Todas las descargas corresponden a la ejecución seleccionada.')
        groups=[('Costos y ahorro',['Comparacion_Fases','Ahorros_Etapas']),
                ('Rutas y vehículos',['FI_Rutas','FI_Visitas','FII_Rutas','FII_Visitas']),
                ('Clientes y bodegas',['FI_Asignaciones','FII_Asignaciones','FI_BodegasAbiertas','FII_BodegasAbiertas']),
                ('Construcciones y mejoras',['FaseI_Lotes','FaseI_Evolucion','Top5_FaseI',
                        'Evaluacion_Candidatos','Movimientos_FaseII','Tiempos_Movimientos','Ciclos_Adicionales']),
                ('Validaciones',['FI_Validacion','FII_Validacion','Control_Ejecucion'])]
        selection=st.selectbox('Grupo de tablas a exportar', [g[0] for g in groups])
        selected_names=next(fields for title,fields in groups if title==selection)
        selected_tables={name:sheet(book,name) for name in selected_names if not sheet(book,name).empty}
        if selected_tables:
            xbuf=BytesIO()
            with pd.ExcelWriter(xbuf,engine='openpyxl') as writer:
                for name,df in selected_tables.items():
                    df.to_excel(writer,sheet_name=name[:31],index=False)
            st.download_button('Descargar tablas en Excel',xbuf.getvalue(),
                 'LogiData_'+re.sub(r'[^A-Za-z]+','_',selection)+'.xlsx',
                 mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            for name,df in selected_tables.items():
                with st.expander(name):
                    data_table(df,use_container_width=True,hide_index=True)
                    download_csv(df,name+'.csv','group_'+name)
        else:
            st.info('No hay registros de este grupo en el Excel seleccionado.')
        st.markdown('#### Informe ejecutivo en PDF')
        if st.button('Preparar informe PDF',key='pdf_prepare'):
            st.session_state['logidata_pdf_data']=crear_informe_pdf()
        if 'logidata_pdf_data' in st.session_state:
            pdf_file=st.session_state['logidata_pdf_data']
            if pdf_file is None:
                st.info('Para PDF instala reportlab y matplotlib. El archivo Excel sigue disponible.')
            else:
                st.download_button('Descargar informe PDF',pdf_file,
                    'LogiData_Informe_Ejecutivo.pdf',mime='application/pdf')
        st.download_button('Descargar Excel consolidado completo',source_bytes,source_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='excel_full_export')

    summary_text=f'''# Informe ejecutivo LogiData\n\nFuente: {source_name}\n\n- Costo diario Fase I: {fmt_money(refi)} COP/día\n- Costo diario Fase II: {fmt_money(refii)} COP/día\n- Ahorro: {fmt_money(saving)} COP/día ({fmt_num(percent,2)} %)\n- Rutas / vehículos finales: {fmt_num(summaries['FII']['rutas'],0)}\n- Bodegas abiertas finales: {fmt_num(summaries['FII']['bodegas'],0)}\n- Distancia final: {fmt_num(summaries['FII']['km'],2)} km\n- Tiempo total de ejecución: {fmt_num(total if 'total' in locals() else np.nan,2)} s\n\nLas conexiones de los mapas son esquemáticas, no geometrías viales. Se muestran únicamente datos disponibles en los registros.\n'''
    st.sidebar.divider()
    st.sidebar.download_button('Descargar resumen ejecutivo',summary_text.encode('utf-8'),
                               'LogiData_Resumen_Ejecutivo.md',mime='text/markdown')
    st.sidebar.download_button('Descargar Excel original',source_bytes,source_name,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


    with tabs[13]:
        st.subheader('Comparación real entre instancias')
        st.caption('Cada fila procede de su propio Excel consolidado. La comparación es descriptiva: las instancias tienen distintas demandas y tamaños.')
        if not archivos_instancias:
            st.info('No hay ejecuciones guardadas en la carpeta de resultados. Utiliza --all-instances para producirlas o indica --results-dir.')
        else:
            filas=[]
            for arch in archivos_instancias:
                try:
                    cmp_row=pd.read_excel(arch,sheet_name='Comparacion_Fases',nrows=1,engine='openpyxl').iloc[0].to_dict()
                    ctrl_row=pd.read_excel(arch,sheet_name='Control_Ejecucion',nrows=1,engine='openpyxl').iloc[0].to_dict()
                    nmatch=re.search(r'Resultados_(\d+)_clientes',arch.name)
                    nclient=int(nmatch.group(1)) if nmatch else None
                    filas.append({'Instancia':arch.stem,'Clientes':nclient,
                       'Costo FI (COP/día)':num(cmp_row.get('costo_inicial')),
                       'Costo FII (COP/día)':num(cmp_row.get('costo_final')),
                       'Costo por cliente (COP/día)':(num(cmp_row.get('costo_final'))/nclient if nclient else np.nan),
                       'Ahorro (COP/día)':num(cmp_row.get('ahorro_cop')),
                       'Reducción (%)':num(cmp_row.get('ahorro_porcentaje')),
                       'Bodegas FI':str(cmp_row.get('bodegas_fase1','')),
                       'Bodegas FII':str(cmp_row.get('bodegas_fase2','')),
                       'Vehículos FI':num(cmp_row.get('vehiculos_fase1')),
                       'Vehículos FII':num(cmp_row.get('vehiculos_fase2')),
                       'Rutas FII':num(cmp_row.get('rutas_fase2',cmp_row.get('vehiculos_fase2'))),
                       'Bodegas abiertas FII':len([x for x in str(cmp_row.get('bodegas_fase2','')).split(' | ') if x]),
                       'km FI':num(cmp_row.get('km_fase1')),
                       'km FII':num(cmp_row.get('km_fase2')),
                       'Tiempo total (s)':num(ctrl_row.get('tiempo_total_segundos')),
                       'Estado': 'Resultado guardado'})
                except Exception as ex:
                    st.warning(f'{arch.name}: no se pudo recuperar la comparación: {ex}')
            if filas:
                comp_inst=pd.DataFrame(filas).sort_values('Clientes')
                data_table(comp_inst,use_container_width=True,hide_index=True)
                download_csv(comp_inst,'LogiData_Comparacion_Instancias.csv','comparacion_multi_csv')
                download_table_xlsx(comp_inst,'LogiData_Comparacion_Instancias.xlsx','comparacion_multi_xlsx')
                for y,title in [('Costo FII (COP/día)','Costo diario final por instancia'),
                                ('Reducción (%)','Reducción porcentual FI → FII'),
                                ('Costo por cliente (COP/día)','Costo diario por cliente'),
                                ('km FII','Kilómetros recorridos'),
                                ('Vehículos FII','Vehículos utilizados'),
                                ('Tiempo total (s)','Tiempo computacional por instancia')]:
                    g=px.bar(comp_inst,x='Clientes',y=y,text_auto='.2f',
                        title=title,color='Clientes',color_continuous_scale='Teal')
                    g.update_layout(height=370,xaxis={'type':'category'},showlegend=False)
                    show_chart(g,use_container_width=True)
                st.markdown('#### Consolidados por instancia')
                for path in archivos_instancias:
                    st.download_button(f'Descargar {path.stem.replace("LogiData_Resultados_","").replace("_"," ")}',
                        path.read_bytes(),file_name=path.name,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=f'download_individual_{path.stem}')
                if st.button('Preparar informe comparativo PDF',key='comparar_pdf_btn'):
                    try:
                        from reportlab.platypus import SimpleDocTemplate,Table,TableStyle,Paragraph,Spacer
                        from reportlab.lib import colors
                        from reportlab.lib.styles import getSampleStyleSheet
                        from reportlab.lib.pagesizes import landscape,A4
                        buff=BytesIO();doc=SimpleDocTemplate(buff,pagesize=landscape(A4))
                        styles=getSampleStyleSheet()
                        headings=['Clientes','Costo FI','Costo FII','Ahorro','Reducción %','km FII','Vehículos FII','Tiempo (s)']
                        fields=['Clientes','Costo FI (COP/día)','Costo FII (COP/día)',
                                'Ahorro (COP/día)','Reducción (%)','km FII','Vehículos FII','Tiempo total (s)']
                        rows=[headings]+[[fmt_num(row.get(k),2) for k in fields]
                                           for _,row in comp_inst.iterrows()]
                        tbl=Table(rows,repeatRows=1,hAlign='LEFT')
                        tbl.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor(NAVY)),
                            ('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTSIZE',(0,0),(-1,-1),8),
                            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F0F5F8')]),
                            ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
                        doc.build([Paragraph('LogiData · Comparación de instancias',styles['Title']),
                            Spacer(1,13),tbl]);st.session_state['comparar_pdf_bytes']=buff.getvalue()
                    except ImportError:
                        st.warning('Instala reportlab para habilitar el informe PDF.')
                if 'comparar_pdf_bytes' in st.session_state:
                    st.download_button('Descargar informe comparativo PDF',
                        st.session_state['comparar_pdf_bytes'],'LogiData_Comparacion_Instancias.pdf',
                        mime='application/pdf',key='comparar_pdf_descarga')


# ============================================================================
# LANZADOR INTEGRADO. No reescribe archivos .py ni ejecuta gráficos al optimizar.
# Uso: py LogiData_Todo_En_Uno.py
# Ver una corrida guardada: py LogiData_Todo_En_Uno.py --view resultados.xlsx
# Solo optimizar: py LogiData_Todo_En_Uno.py --no-dashboard
# Seleccionar cuatro Excels: py LogiData_Todo_En_Uno.py --select-files
# Dependencias: py -m pip install pandas numpy scipy openpyxl matplotlib streamlit plotly reportlab
# ============================================================================


def _logidata_pick_inputs():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as exc:
        raise RuntimeError('Tkinter no disponible: pasa --bodegas --clientes --distancias --tiempos.') from exc
    w = tk.Tk()
    w.withdraw()
    try:
        result=[]
        for param,descr in [('bodegas','bodegas'),('clientes','clientes'),
                            ('distancias','distancias'),('tiempos','tiempos')]:
            p=filedialog.askopenfilename(title='LogiData: Excel de '+descr,
                   initialdir=str(_logidata_root()),filetypes=[('Excel','*.xlsx')])
            if not p:
                raise RuntimeError('Selección cancelada. No se ejecutó la optimización.')
            result.extend(['--'+param,p])
        return result
    finally:
        w.destroy()


def _logidata_root():
    from pathlib import Path
    return Path(__file__).resolve().parent


def _logidata_auto_inputs():
    names={
      'bodegas':('Localización de Instalaciones Proyecto.xlsx',
                 'Localización de Instalaciones Proyecto(6).xlsx'),
      'clientes':('Instancia_01 (1).xlsx','Instancia_01 (1)(8).xlsx'),
      'distancias':('Distancias_LogiData(1).xlsx','Distancias_LogiData (1)(1).xlsx'),
      'tiempos':('Tiempos_LogiData(1).xlsx','Tiempos_LogiData (1)(1).xlsx')}
    found=[]
    for param,options in names.items():
        path=next((_logidata_root()/name for name in options if
                   (_logidata_root()/name).is_file()),None)
        if path is not None: found.extend(['--'+param,str(path)])
    return found


def _logidata_port_occupied(port):
    import socket
    try:
        with socket.create_connection(('127.0.0.1',port),timeout=.25):
            return True
    except OSError:
        return False


def _logidata_start_dashboard(excel, port, results_dir=None):
    from pathlib import Path
    import subprocess, sys, time, webbrowser
    excel=Path(excel).expanduser().resolve()
    if not excel.is_file():
        raise FileNotFoundError('No se encuentra el Excel consolidado: '+str(excel))
    try:
        import importlib.util
        if importlib.util.find_spec('streamlit') is None:
            raise RuntimeError('Streamlit no instalado. Ejecuta: py -m pip install streamlit plotly')
    except ImportError as exc:
        raise RuntimeError('No se pudo comprobar Streamlit') from exc
    if _logidata_port_occupied(port):
        port=next((p for p in range(port+1,min(port+51,65536))
                   if not _logidata_port_occupied(p)),None)
        if port is None: raise RuntimeError('No se encontró puerto libre para Streamlit')
    # Una instancia nueva lee el Excel exacto de ESTA ejecución; ningún dato queda precargado.
    dashboard_args=['--dashboard','--excel',str(excel)]
    if results_dir is not None:
        dashboard_args+=['--results-dir',str(Path(results_dir).resolve())]
    proc=subprocess.Popen([sys.executable,'-m','streamlit','run',
            str(Path(__file__).resolve()),'--server.port',str(port),
            '--server.headless','true','--',*dashboard_args],
            cwd=str(_logidata_root()))
    url='http://localhost:'+str(port)
    try:
        for _ in range(120):
            if proc.poll() is not None:
                raise RuntimeError('Streamlit finalizó antes de abrir. Revisa dependencias y consola.')
            if _logidata_port_occupied(port):
                webbrowser.open(url)
                print('DASHBOARD:',url,flush=True)
                print('Excel analizado:',excel,flush=True)
                break
            time.sleep(.10)
        else:
            print('Streamlit continúa iniciando. Abre manualmente:',url,flush=True)
        return proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        try: proc.wait(timeout=5)
        except subprocess.TimeoutExpired: proc.kill()
        return 0
    except Exception:
        proc.terminate()
        raise



def _logidata_catalogo_instancias(root=None):
    """Encuentra únicamente instancias de entrada, nunca Excel consolidados."""
    import re
    from pathlib import Path
    base=Path(root).expanduser().resolve() if root else _logidata_root()
    encontrados={}
    for archivo in sorted(base.glob('LogiData_Instancia_*_clientes*.xlsx')):
        m=re.search(r'Instancia_(\d+)_(\d+)_clientes',archivo.name,re.I)
        if not m: continue
        clientes=int(m.group(2))
        if clientes not in (5,10,15,20,25): continue
        if clientes in encontrados:
            raise ValueError(f'Se encontraron dos Excel de {clientes} clientes en {base}. Deja solo uno por tamaño o usa --instance con la ruta exacta.')
        encontrados[clientes]=archivo
    return dict(sorted(encontrados.items()))


def _logidata_resolver_instancia(value):
    """Acepta número de clientes (5/10/15/20/25) o ruta .xlsx."""
    from pathlib import Path
    if str(value).isdigit():
        opciones=_logidata_catalogo_instancias()
        numero=int(value)
        if numero not in opciones:
            raise FileNotFoundError(f'No existe un archivo de instancia para {numero} clientes junto al programa.')
        return opciones[numero],numero
    path=Path(value).expanduser().resolve()
    if not path.is_file() or path.suffix.lower()!='.xlsx':
        raise FileNotFoundError(f'Instancia Excel no encontrada: {value}')
    # Leer SOLO metadatos para asegurar que el nombre no determina el tamaño.
    import pandas as pd
    cantidad=len(pd.read_excel(path,sheet_name='Sheet1',usecols=[0],engine='openpyxl'))
    return path,cantidad


def _logidata_ejecutar_instancia(path,clients,target,extra):
    from pathlib import Path
    path=Path(path).resolve();target=Path(target).resolve()
    target.parent.mkdir(parents=True,exist_ok=True)
    entradas=[item for key in ('bodegas','clientes','distancias','tiempos')
              for item in ('--'+key,str(path))]
    print('\n'+'='*65+f'\nINSTANCIA: {clients} clientes | {path.name}\n'+'='*65,flush=True)
    ejecutar_integrado([*entradas,*extra,'--output',str(target)])
    if not target.is_file():
        raise RuntimeError(f'No se generó el Excel para {clients} clientes')
    print('Consolidado de instancia:',target,flush=True)
    return target

def _logidata_entrada(argv=None):
    import argparse
    import sys
    from pathlib import Path
    args=list(sys.argv[1:] if argv is None else argv)
    # Streamlit (con o sin --dashboard) NUNCA ejecuta la heurística al rerenderizar.
    # Comprobación perezosa: no importar Streamlit durante la optimización normal.
    en_streamlit = False
    if 'streamlit' in sys.modules:
        try:
            from streamlit.runtime.scriptrunner import get_script_run_ctx
            en_streamlit = get_script_run_ctx(suppress_warning=True) is not None
        except (ImportError, TypeError):
            pass
    if '--dashboard' in args or en_streamlit:
        return _logidata_dashboard()
    ap=argparse.ArgumentParser(description=(
        'LogiData TODO EN UNO: ejecuta Fase I y II, genera Excel y abre Streamlit. '
        'Los parámetros desconocidos se reenvían al algoritmo original.'))
    ap.add_argument('--view',metavar='RESULTADOS.xlsx',help='Visualizar una ejecución guardada, sin optimizar')
    ap.add_argument('--select-files',action='store_true',help='Elegir los cuatro Excel originales en ventanas')
    ap.add_argument('--output',default=str(_logidata_root()/'LogiData_Fases_I_II_consolidado.xlsx'))
    ap.add_argument('--port',type=int,default=8501)
    ap.add_argument('--no-dashboard',action='store_true',help='Solo optimizar y guardar Excel')
    ap.add_argument('--instance',help='Instancia individual: 5, 10, 15, 20, 25 o ruta del Excel único')
    ap.add_argument('--all-instances',action='store_true',help='Resolver las cinco instancias consecutivamente en una sola ejecución y abrir el dashboard')
    ap.add_argument('--instances-dir',help='Carpeta con las cinco instancias (por defecto, junto al .py)')
    ap.add_argument('--results-dir',default=str(_logidata_root()/'resultados_logidata'),
                    help='Carpeta para los consolidados independientes')
    ap.add_argument('--help-optimizer',action='store_true',help='Ver opciones originales de Fases I y II')
    opts, extra=ap.parse_known_args(args)
    if opts.help_optimizer:
        return ejecutar_integrado(['--help'])
    if not 1 <= opts.port <= 65535:
        ap.error('Puerto inválido')
    if opts.view:
        return _logidata_start_dashboard(opts.view,opts.port,opts.results_dir)
    if opts.instance and opts.all_instances:
        ap.error('Selecciona --instance o --all-instances, no ambos.')
    target=Path(opts.output).expanduser().resolve()
    if target.suffix.lower()!='.xlsx':
        ap.error('--output debe terminar en .xlsx')
    # Verificar dependencias gráficas al principio para no optimizar durante horas
    # y descubrir al final que falta Streamlit.
    if not opts.no_dashboard:
        import importlib.util
        missing=[x for x in ('streamlit','plotly') if importlib.util.find_spec(x) is None]
        if missing:
            raise RuntimeError('Instala las dependencias del dashboard: '
                 'py -m pip install streamlit plotly (faltan: '+', '.join(missing)+')')
    # Modo multiinstancia, sin cambiar parámetros, semillas ni rutinas heurísticas.
    if opts.instance or opts.all_instances:
        if opts.all_instances:
            catalogo=_logidata_catalogo_instancias(opts.instances_dir)
            if not catalogo:
                raise FileNotFoundError('No hay instancias junto al programa. Usa --instances-dir CARPETA.')
            missing = sorted({5,10,15,20,25}-set(catalogo))
            if missing:
                raise FileNotFoundError('Faltan instancias de '+', '.join(str(n) for n in missing)+
                    ' clientes. Coloca los cinco archivos de entrada en la misma carpeta antes de ejecutar --all-instances.')
            targets=list(catalogo.items())
        elif opts.instance.isdigit() and opts.instances_dir:
            catalogo=_logidata_catalogo_instancias(opts.instances_dir)
            n=int(opts.instance)
            if n not in catalogo: raise FileNotFoundError(f'No hay instancia de {n} clientes en {opts.instances_dir}')
            targets=[(n,catalogo[n])]
        else:
            path,n=_logidata_resolver_instancia(opts.instance)
            targets=[(n,path)]
        dest=Path(opts.results_dir).expanduser().resolve()
        dest.mkdir(parents=True,exist_ok=True)
        productos=[];errores=[]
        for n,instancia in targets:
            result_path=dest/f'LogiData_Resultados_{n:02d}_clientes.xlsx'
            try:
                productos.append(_logidata_ejecutar_instancia(instancia,n,result_path,extra))
            except Exception as ex:
                print(f'ERROR de instancia {n}: {ex}',file=sys.stderr,flush=True)
                errores.append((n,str(ex)))
                if not opts.all_instances: raise
        if not productos:
            raise RuntimeError('Ninguna instancia generó resultados correctos: '+str(errores))
        print(f'Ejecuciones terminadas: {len(productos)} de {len(targets)}',flush=True)
        if errores: print('Instancias con error:',errores,file=sys.stderr)
        if opts.no_dashboard: return 1 if errores else 0
        status=_logidata_start_dashboard(productos[0],opts.port,dest)
        return 1 if errores else status
    inputs=_logidata_pick_inputs() if opts.select_files else _logidata_auto_inputs()
    target.parent.mkdir(parents=True,exist_ok=True)
    print('Iniciando Fases I y II originales, sin gráficos durante la optimización.',flush=True)
    print('Archivo de salida:',target,flush=True)
    # Única ejecución de la heurística en este proceso, sin cambiar algoritmos ni parámetros.
    ejecutar_integrado([*inputs,*extra,'--output',str(target)])
    if not target.is_file():
        raise RuntimeError('La heurística terminó sin producir el Excel consolidado.')
    print('Excel consolidado creado correctamente:',target,flush=True)
    if opts.no_dashboard:
        return 0
    return _logidata_start_dashboard(target,opts.port,opts.results_dir)


if __name__=='__main__':
    try:
        import sys
        sys.exit(_logidata_entrada())
    except (RuntimeError,FileNotFoundError,ValueError) as exc:
        import sys
        print('ERROR:',exc,file=sys.stderr)
        sys.exit(1)
