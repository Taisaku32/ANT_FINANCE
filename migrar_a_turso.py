"""Migra los datos locales de un usuario (users/<usuario>/finanzas.db) a una
base de datos Turso en la nube.

Uso:
    python migrar_a_turso.py --usuario "juan manuel" --url libsql://xxx.turso.io --token eyJ...

La base local se abre en modo SOLO LECTURA: este script nunca modifica ni
borra los archivos locales. La base destino en Turso debe estar vacía.
"""

import argparse
import os
import sqlite3
import sys

from turso_db import TursoConnection

# Esquema canónico (el mismo de main.py + dashboard.py)
DDL = [
    '''CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)''',
    '''CREATE TABLE IF NOT EXISTS subcategories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id INTEGER NOT NULL REFERENCES categories(id),
        name TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)''',
    '''CREATE TABLE IF NOT EXISTS ingresos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        monto REAL, categoria TEXT, fecha TEXT,
        category_id INTEGER, subcategory_id INTEGER, activity_name TEXT)''',
    '''CREATE TABLE IF NOT EXISTS gastos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        monto REAL, categoria TEXT, fecha TEXT,
        category_id INTEGER, subcategory_id INTEGER, activity_name TEXT)''',
    '''CREATE TABLE IF NOT EXISTS ahorros (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        monto          REAL    NOT NULL,
        tipo           TEXT    NOT NULL CHECK(tipo IN ('deposito', 'retiro')),
        categoria      TEXT,
        category_id    INTEGER REFERENCES categories(id),
        subcategory_id INTEGER REFERENCES subcategories(id),
        activity_name  TEXT,
        fecha          TEXT    NOT NULL)''',
    '''CREATE TABLE IF NOT EXISTS budgets (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        categoria TEXT    NOT NULL UNIQUE,
        monto     REAL    NOT NULL CHECK(monto > 0))''',
    '''CREATE TABLE IF NOT EXISTS balances_mensuales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        año INTEGER, mes INTEGER, balance REAL, fecha_creacion TEXT,
        UNIQUE(año, mes))''',
    '''CREATE TABLE IF NOT EXISTS metas_ahorro (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id    INTEGER NOT NULL REFERENCES categories(id),
        subcategory_id INTEGER REFERENCES subcategories(id),
        nombre         TEXT    NOT NULL,
        monto_objetivo REAL    NOT NULL CHECK(monto_objetivo > 0),
        created_at     TEXT    DEFAULT CURRENT_TIMESTAMP)''',
]

# Orden que respeta las referencias entre tablas
TABLAS = ['categories', 'subcategories', 'ingresos', 'gastos', 'ahorros',
          'budgets', 'balances_mensuales', 'metas_ahorro']


def main():
    parser = argparse.ArgumentParser(description='Migrar datos locales a Turso.')
    parser.add_argument('--usuario', required=True,
                        help='Nombre de la carpeta en users/ (ej: "juan manuel")')
    parser.add_argument('--url', required=True,
                        help='URL de la base de datos Turso (libsql://...)')
    parser.add_argument('--token', required=True, help='Token de autenticación de Turso')
    parser.add_argument('--lote', type=int, default=50,
                        help='Filas por petición HTTP (por defecto 50)')
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(base_dir, 'users', args.usuario, 'finanzas.db')
    if not os.path.exists(db_path):
        users_dir = os.path.join(base_dir, 'users')
        disponibles = sorted(os.listdir(users_dir)) if os.path.isdir(users_dir) else []
        sys.exit(f'No existe {db_path}\nUsuarios disponibles: {disponibles}')

    # Base local en SOLO LECTURA
    local = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    remoto = TursoConnection(args.url, args.token)

    print(f'Origen : {db_path}')
    print(f'Destino: {args.url}')

    # 1. Crear el esquema en Turso
    for ddl in DDL:
        remoto.execute(ddl)
    print('Esquema creado en Turso.')

    # 2. Verificar que el destino esté vacío (para no duplicar datos)
    for t in TABLAS:
        n = remoto.execute(f'SELECT COUNT(*) FROM {t}').fetchone()[0]
        if n:
            sys.exit(f'La tabla "{t}" en Turso ya tiene {n} filas. '
                     'Usa una base de datos vacía o borra su contenido desde el panel de Turso.')

    # 3. Copiar los datos tabla por tabla
    total = 0
    for t in TABLAS:
        existe = local.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (t,)
        ).fetchone()
        if not existe:
            print(f'{t}: no existe en la base local, se omite.')
            continue
        cols = [r[1] for r in local.execute(f'PRAGMA table_info({t})').fetchall()]
        filas = local.execute(f'SELECT {", ".join(cols)} FROM {t}').fetchall()
        if not filas:
            print(f'{t}: 0 filas.')
            continue
        marcadores = ', '.join('?' * len(cols))
        sql = f'INSERT INTO {t} ({", ".join(cols)}) VALUES ({marcadores})'
        for i in range(0, len(filas), args.lote):
            remoto.execute_batch([(sql, tuple(f)) for f in filas[i:i + args.lote]])
        print(f'{t}: {len(filas)} filas migradas.')
        total += len(filas)

    local.close()
    print(f'\nMigración completa: {total} filas copiadas a Turso.')
    print('Los archivos locales NO fueron modificados.')


if __name__ == '__main__':
    main()
