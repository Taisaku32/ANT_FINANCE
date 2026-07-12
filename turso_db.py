"""Cliente mínimo para Turso (libSQL) sobre HTTP (protocolo Hrana v2, endpoint /v2/pipeline).

No requiere dependencias nativas: solo `requests`. Expone una interfaz compatible
con el módulo `sqlite3` (execute / cursor / commit / close) para que dashboard.py
pueda usar la misma lógica tanto en modo local (SQLite) como en la nube (Turso).

Cada `execute()` se auto-confirma en el servidor (no hay transacciones
interactivas); `commit()` y `close()` existen solo por compatibilidad.
"""

import base64
import math

import requests


def _a_arg_hrana(valor):
    """Convierte un valor Python al formato de argumento de Hrana."""
    # Escalares de numpy/pandas (int64, float64, ...) → escalar Python
    if hasattr(valor, "item") and not isinstance(valor, (str, bytes, bytearray)):
        try:
            valor = valor.item()
        except Exception:
            pass
    if valor is None:
        return {"type": "null"}
    if isinstance(valor, bool):
        return {"type": "integer", "value": str(int(valor))}
    if isinstance(valor, int):
        return {"type": "integer", "value": str(valor)}
    if isinstance(valor, float):
        if math.isnan(valor):
            return {"type": "null"}
        return {"type": "float", "value": valor}
    if isinstance(valor, (bytes, bytearray)):
        return {"type": "blob", "base64": base64.b64encode(bytes(valor)).decode("ascii")}
    return {"type": "text", "value": str(valor)}


def _de_valor_hrana(celda):
    """Convierte una celda de respuesta de Hrana a un valor Python."""
    tipo = celda.get("type")
    if tipo == "null":
        return None
    if tipo == "integer":
        return int(celda["value"])
    if tipo == "float":
        return float(celda["value"])
    if tipo == "blob":
        return base64.b64decode(celda.get("base64", "") + "==")
    return celda.get("value")


class TursoCursor:
    def __init__(self, conn):
        self._conn = conn
        self.description = None
        self.rowcount = -1
        self.lastrowid = None
        self._rows = []
        self._idx = 0

    def execute(self, sql, params=()):
        resultado = self._conn._ejecutar(sql, params)
        cols = resultado.get("cols") or []
        self.description = (
            [(c.get("name"), None, None, None, None, None, None) for c in cols] or None
        )
        self._rows = [
            tuple(_de_valor_hrana(v) for v in fila)
            for fila in (resultado.get("rows") or [])
        ]
        self._idx = 0
        rc = resultado.get("affected_row_count")
        self.rowcount = rc if rc is not None else -1
        last = resultado.get("last_insert_rowid")
        self.lastrowid = int(last) if last is not None else None
        return self

    def fetchall(self):
        filas = self._rows[self._idx:]
        self._idx = len(self._rows)
        return filas

    def fetchone(self):
        if self._idx < len(self._rows):
            fila = self._rows[self._idx]
            self._idx += 1
            return fila
        return None

    def close(self):
        pass


class TursoConnection:
    """Conexión a una base de datos Turso vía HTTP."""

    def __init__(self, url, auth_token, timeout=30):
        base = (url or "").strip().rstrip("/")
        for esquema in ("libsql://", "wss://", "ws://", "http://"):
            if base.startswith(esquema):
                base = "https://" + base[len(esquema):]
                break
        if not base.startswith("https://"):
            base = "https://" + base
        self._endpoint = base + "/v2/pipeline"
        self._headers = {
            "Authorization": "Bearer " + (auth_token or "").strip(),
            "Content-Type": "application/json",
        }
        self._timeout = timeout

    def _pipeline(self, stmts):
        cuerpo = [{"type": "execute", "stmt": s} for s in stmts]
        cuerpo.append({"type": "close"})
        resp = requests.post(
            self._endpoint,
            json={"requests": cuerpo},
            headers=self._headers,
            timeout=self._timeout,
        )
        resp.raise_for_status()
        data = resp.json()
        resultados = []
        for res in data.get("results", []):
            if res.get("type") == "error":
                mensaje = (res.get("error") or {}).get("message", "error desconocido")
                raise RuntimeError(f"Error de Turso: {mensaje}")
            respuesta = res.get("response") or {}
            if respuesta.get("type") == "execute":
                resultados.append(respuesta.get("result") or {})
        return resultados

    def _ejecutar(self, sql, params=()):
        stmt = {"sql": sql}
        if params:
            stmt["args"] = [_a_arg_hrana(p) for p in params]
        return self._pipeline([stmt])[0]

    def execute(self, sql, params=()):
        return TursoCursor(self).execute(sql, params)

    def execute_batch(self, sentencias):
        """Ejecuta varias sentencias [(sql, params), ...] en una sola petición HTTP."""
        stmts = []
        for sql, params in sentencias:
            stmt = {"sql": sql}
            if params:
                stmt["args"] = [_a_arg_hrana(p) for p in params]
            stmts.append(stmt)
        return self._pipeline(stmts)

    def cursor(self):
        return TursoCursor(self)

    def commit(self):
        pass

    def close(self):
        pass
