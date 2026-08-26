3#!/usr/bin/env python3
"""Consolida JSONs do machine_scanner em dados/inventario.csv (SATO-xxx / MON-xxx)."""

from __future__ import annotations

import argparse
import csv
import json
import re
import traceback
from datetime import date, datetime
from pathlib import Path
from typing import Any

CSV_COLUMNS = [
    "Codigo",
    "Tipo",
    "SerialNumber",
    "Hostname",
    "Usuario",
    "Fabricante",
    "Modelo",
    "CPU",
    "RAM_GB",
    "Disco_GB",
    "SO",
    "MAC",
    "VinculadoA",
    "DataColeta",
    "ArquivoOrigem",
    "Observacao",
]

CODE_RE = re.compile(r"^(SATO|MON)-(\d+)$", re.IGNORECASE)


def _root() -> Path:
    return Path(__file__).resolve().parent


def _section_map(payload: dict[str, Any]) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for section in payload.get("sections") or []:
        if not isinstance(section, dict):
            continue
        name = section.get("name")
        if name:
            out[str(name)] = section
    return out


def _data(section: dict[str, Any] | None) -> dict[str, Any]:
    if not section:
        return {}
    data = section.get("data")
    return data if isinstance(data, dict) else {}


_BAD_SERIALS = {
    "",
    "0",
    "none",
    "null",
    "n/a",
    "unknown",
    "-",
    "to be filled by o.e.m.",
    "default string",
    "system serial number",
}


def _clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"none", "null", "n/a", "unknown", "-"}:
        return ""
    return text


def _clean_serial(value: Any) -> str:
    text = _clean(value)
    if text.casefold() in _BAD_SERIALS:
        return ""
    # Drop control chars sometimes present in USB device serials.
    text = "".join(ch for ch in text if ch.isprintable())
    return text.strip()


def _pick_serial(baseboard: dict[str, Any]) -> str:
    for key in ("system_serial", "baseboard_serial", "bios_serial", "chassis_serial"):
        serial = _clean_serial(baseboard.get(key))
        if serial:
            return serial
    return ""


def _mac_from_iface(iface: dict[str, Any]) -> str:
    name = _clean(iface.get("name")).casefold()
    if "loopback" in name or "bluetooth" in name:
        return ""
    for addr in iface.get("addresses") or []:
        if not isinstance(addr, dict):
            continue
        family = str(addr.get("family") or "")
        address = _clean(addr.get("address"))
        if not address:
            continue
        if "LINK" in family.upper() or "PACKET" in family.upper() or family == "17":
            if address.lower() in {"00:00:00:00:00:00", "00-00-00-00-00-00"}:
                continue
            return address.replace("-", ":").upper()
    return ""


def _first_mac(network: dict[str, Any]) -> str:
    ifaces = [i for i in (network.get("interfaces") or []) if isinstance(i, dict)]
    # Prefer interfaces that are up; if Wi-Fi is down, still keep a physical MAC.
    for iface in ifaces:
        if iface.get("is_up") is True:
            mac = _mac_from_iface(iface)
            if mac:
                return mac
    for iface in ifaces:
        mac = _mac_from_iface(iface)
        if mac:
            return mac
    return ""


def _disk_gb(sections: dict[str, dict[str, Any]]) -> str:
    storage = _data(sections.get("storage_devices"))
    drives = storage.get("drives") or []
    total = 0.0
    found = False
    for drive in drives:
        if not isinstance(drive, dict):
            continue
        bus = _clean(drive.get("bus")).upper()
        model = _clean(drive.get("model")).casefold()
        # Ignore the inventory USB stick itself.
        if bus == "USB" or "datatraveler" in model or "kingston" in model:
            continue
        size = drive.get("size_gb")
        if size is None:
            continue
        try:
            total += float(size)
            found = True
        except (TypeError, ValueError):
            pass
    if found:
        return str(int(round(total)))

    # Fallback: system drive C: only (not the pendrive letter).
    disk = _data(sections.get("disk"))
    for part in disk.get("partitions") or []:
        if not isinstance(part, dict):
            continue
        mount = _clean(part.get("mountpoint") or part.get("device")).upper()
        if mount not in {"C:\\", "C:"}:
            continue
        try:
            return str(int(round(float(part.get("total_gb") or 0))))
        except (TypeError, ValueError):
            return ""
    return ""


def _machine_tipo(sections: dict[str, dict[str, Any]]) -> str:
    battery = _data(sections.get("battery"))
    # Presence of battery entries usually means notebook/laptop.
    if battery.get("present") is True:
        return "Notebook"
    batteries = battery.get("batteries") or battery.get("devices") or []
    if isinstance(batteries, list) and batteries:
        return "Notebook"
    status = (sections.get("battery") or {}).get("status")
    if status == "ok":
        return "Notebook"
    return "Desktop"


_GENERIC_INPUT = {
    "aperfeiçoado (101 ou 102 teclas)",
    "aperfeicoado (101 ou 102 teclas)",
    "standard ps/2 keyboard",
    "hid keyboard device",
    "mouse compatível com hid",
    "mouse compativel com hid",
    "hid-compliant mouse",
}


def _input_obs(sections: dict[str, dict[str, Any]]) -> str:
    devices = _data(sections.get("input")).get("devices") or []
    names: list[str] = []
    for device in devices:
        if not isinstance(device, dict):
            continue
        kind = _clean(device.get("kind") or device.get("type")).lower()
        name = _clean(device.get("name"))
        if not name:
            continue
        if name.casefold() in _GENERIC_INPUT:
            continue
        if "point" in kind or "mouse" in kind:
            names.append(name)
    # Keep short; avoid huge HID noise.
    names = list(dict.fromkeys(names))[:4]
    return "; ".join(names)


def _so_text(meta: dict[str, Any], system: dict[str, Any]) -> str:
    detail = _clean(meta.get("os_detail"))
    if detail:
        return detail
    parts = [
        _clean(system.get("system")),
        _clean(system.get("release")),
    ]
    return " ".join(p for p in parts if p)


def parse_scan(path: Path) -> tuple[dict[str, str], list[dict[str, str]]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    meta = payload.get("meta") if isinstance(payload.get("meta"), dict) else {}
    sections = _section_map(payload)

    baseboard = _data(sections.get("baseboard"))
    network = _data(sections.get("network"))
    system = _data(sections.get("system"))
    cpu = _data(sections.get("cpu"))
    memory = _data(sections.get("memory"))

    serial = _pick_serial(baseboard)
    obs_parts: list[str] = []
    if not serial:
        obs_parts.append("SEM_SERIAL")
    input_obs = _input_obs(sections)
    if input_obs:
        obs_parts.append(input_obs)

    machine = {
        "Tipo": _machine_tipo(sections),
        "SerialNumber": serial,
        "Hostname": _clean(meta.get("hostname") or network.get("hostname")),
        "Usuario": _clean(meta.get("user")),
        "Fabricante": _clean(baseboard.get("system_manufacturer") or baseboard.get("baseboard_manufacturer")),
        "Modelo": _clean(baseboard.get("system_product")),
        "CPU": _clean(cpu.get("name") or system.get("processor")),
        "RAM_GB": _clean(memory.get("total_gb")),
        "Disco_GB": _disk_gb(sections),
        "SO": _so_text(meta, system),
        "MAC": _first_mac(network),
        "VinculadoA": "",
        "DataColeta": _clean(meta.get("scanned_at"))[:10] or date.today().isoformat(),
        "ArquivoOrigem": path.name,
        "Observacao": " | ".join(obs_parts),
    }

    monitors: list[dict[str, str]] = []
    for monitor in _data(sections.get("monitors")).get("monitors") or []:
        if not isinstance(monitor, dict):
            continue
        mon_serial = _clean_serial(monitor.get("serial"))
        mon_model = _clean(monitor.get("name") or monitor.get("product_code"))
        mon_mfr = _clean(monitor.get("manufacturer"))
        if not mon_serial and not mon_model and not mon_mfr:
            continue
        monitors.append(
            {
                "Tipo": "Monitor",
                "SerialNumber": mon_serial,
                "Hostname": machine["Hostname"],
                "Usuario": machine["Usuario"],
                "Fabricante": mon_mfr,
                "Modelo": mon_model,
                "CPU": "",
                "RAM_GB": "",
                "Disco_GB": "",
                "SO": "",
                "MAC": "",
                "VinculadoA": "",  # filled after machine code assigned
                "DataColeta": machine["DataColeta"],
                "ArquivoOrigem": path.name,
                "Observacao": "" if mon_serial else "SEM_SERIAL",
            }
        )

    return machine, monitors


def load_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        rows = []
        for row in reader:
            rows.append({col: row.get(col, "") or "" for col in CSV_COLUMNS})
        return rows


def save_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in CSV_COLUMNS})


def next_code(rows: list[dict[str, str]], prefix: str) -> str:
    max_n = 0
    for row in rows:
        match = CODE_RE.match(_clean(row.get("Codigo")))
        if not match:
            continue
        if match.group(1).upper() != prefix.upper():
            continue
        max_n = max(max_n, int(match.group(2)))
    return f"{prefix.upper()}-{max_n + 1:03d}"


def find_by_serial(rows: list[dict[str, str]], serial: str, tipo_prefix: str | None = None) -> int | None:
    serial_norm = serial.strip().casefold()
    if not serial_norm:
        return None
    for idx, row in enumerate(rows):
        if _clean(row.get("SerialNumber")).casefold() != serial_norm:
            continue
        codigo = _clean(row.get("Codigo")).upper()
        if tipo_prefix and not codigo.startswith(tipo_prefix.upper() + "-"):
            continue
        return idx
    return None


def upsert_machine(rows: list[dict[str, str]], machine: dict[str, str]) -> str:
    serial = machine["SerialNumber"]
    idx = find_by_serial(rows, serial, "SATO") if serial else None
    if idx is None and not serial:
        # Without serial, match by hostname+arquivo is unsafe for updates;
        # always create a new code and mark SEM_SERIAL.
        pass
    elif idx is not None:
        codigo = rows[idx]["Codigo"]
        updated = dict(rows[idx])
        updated.update(machine)
        updated["Codigo"] = codigo
        rows[idx] = updated
        return codigo

    codigo = next_code(rows, "SATO")
    row = dict(machine)
    row["Codigo"] = codigo
    rows.append(row)
    return codigo


def upsert_monitor(rows: list[dict[str, str]], monitor: dict[str, str], sato_code: str) -> str:
    monitor = dict(monitor)
    monitor["VinculadoA"] = sato_code
    serial = monitor["SerialNumber"]
    idx = find_by_serial(rows, serial, "MON") if serial else None
    if idx is not None:
        codigo = rows[idx]["Codigo"]
        updated = dict(rows[idx])
        updated.update(monitor)
        updated["Codigo"] = codigo
        rows[idx] = updated
        return codigo

    codigo = next_code(rows, "MON")
    monitor["Codigo"] = codigo
    rows.append(monitor)
    return codigo


class _Logger:
    def __init__(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self._fh = path.open("a", encoding="utf-8")

    def log(self, message: str) -> None:
        line = f"[{datetime.now().isoformat(timespec='seconds')}] {message}"
        print(message)
        self._fh.write(line + "\n")
        self._fh.flush()

    def close(self) -> None:
        self._fh.close()


def consolidate(scans_dir: Path, csv_path: Path, logger: _Logger) -> tuple[int, int, int]:
    files = sorted(scans_dir.glob("*.json"))
    if not files:
        raise SystemExit(f"Nenhum JSON em {scans_dir}")

    logger.log(f"Scans: {scans_dir} ({len(files)} arquivo(s))")
    logger.log(f"CSV destino: {csv_path}")

    rows = load_csv(csv_path)
    machines = 0
    monitors = 0
    errors = 0

    for path in files:
        try:
            machine, mons = parse_scan(path)
            sato = upsert_machine(rows, machine)
            machines += 1
            logger.log(
                f"OK maquina {path.name} -> {sato} serial={machine.get('SerialNumber') or 'SEM_SERIAL'}"
            )
            for mon in mons:
                mon_code = upsert_monitor(rows, mon, sato)
                monitors += 1
                logger.log(
                    f"OK monitor {path.name} -> {mon_code} serial={mon.get('SerialNumber') or 'SEM_SERIAL'} vinculado={sato}"
                )
        except Exception as exc:
            errors += 1
            logger.log(f"ERRO em {path.name}: {type(exc).__name__}: {exc}")
            logger.log(traceback.format_exc())

    if machines == 0 and errors:
        raise SystemExit(f"Nenhuma maquina consolidada. Veja o log: {logger.path}")

    save_csv(csv_path, rows)
    logger.log(f"CSV gravado com {len(rows)} linha(s)")
    return machines, monitors, errors


def main() -> None:
    parser = argparse.ArgumentParser(description="Consolida scans JSON em inventario.csv")
    parser.add_argument(
        "--scans",
        type=Path,
        default=_root() / "dados" / "scans",
        help="Pasta com os JSONs do machine_scanner",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=_root() / "dados" / "inventario.csv",
        help="CSV de saida",
    )
    parser.add_argument(
        "--log",
        type=Path,
        default=None,
        help="Arquivo de log (padrao: dados/logs/consolidacao_DATAHORA.log)",
    )
    args = parser.parse_args()

    log_path = args.log or (
        _root() / "dados" / "logs" / f"consolidacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    logger = _Logger(log_path)
    try:
        logger.log("=== consolidacao inicio ===")
        machines, monitors, errors = consolidate(args.scans, args.csv, logger)
        logger.log(
            f"Processados: {machines} maquina(s), {monitors} monitor(es), {errors} erro(s)"
        )
        logger.log(f"CSV: {args.csv}")
        logger.log(f"Log: {log_path}")
        logger.log("=== consolidacao fim ===")
    except SystemExit as exc:
        logger.log(f"ERRO FATAL: {exc}")
        raise
    except Exception as exc:
        logger.log(f"ERRO FATAL: {type(exc).__name__}: {exc}")
        logger.log(traceback.format_exc())
        raise SystemExit(1) from exc
    finally:
        logger.close()


if __name__ == "__main__":
    main()
