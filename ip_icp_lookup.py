"""
IP to domain and ICP lookup utility.

Reads a text file containing one IPv4 address per line, looks up the reverse DNS
hostname, retrieves ICP/filing information for the resolved domain, and writes
results to a timestamped Excel file.
"""
from __future__ import annotations

import argparse
import datetime as dt
import ipaddress
import socket
from dataclasses import dataclass
from typing import Iterable, List, Optional

import openpyxl
import requests


@dataclass
class LookupResult:
    ip: str
    domain: str | None
    icp_subject: str | None
    icp_license: str | None
    message: str


def read_ip_file(path: str) -> List[str]:
    with open(path, "r", encoding="utf-8") as handle:
        ips = []
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                ipaddress.ip_address(line)
                ips.append(line)
            except ValueError:
                print(f"Skipping invalid IP entry: {line}")
        return ips


def reverse_lookup(ip: str, session: requests.Session, token: str | None) -> tuple[Optional[str], str]:
    # Attempt socket PTR lookup first.
    try:
        host, _, _ = socket.gethostbyaddr(ip)
        return host, "Resolved via PTR"
    except Exception as exc:
        ptr_error = f"PTR lookup failed: {exc}"

    # Fallback to ipinfo.io, which often provides a hostname field.
    url = f"https://ipinfo.io/{ip}/json"
    try:
        headers = {}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        resp = session.get(url, timeout=10, headers=headers)
        if resp.status_code == 200:
            hostname = resp.json().get("hostname")
            if hostname:
                return hostname, "Resolved via ipinfo.io"
            return None, "ipinfo.io returned no hostname"
        return None, f"ipinfo.io status {resp.status_code}"
    except Exception as exc:  # noqa: BLE001
        return None, f"{ptr_error}; ipinfo.io lookup failed: {exc}"

    return None, ptr_error


def query_icp(domain: str, session: requests.Session) -> tuple[Optional[str], Optional[str], str]:
    url = "https://api.vore.top/api/icp"
    try:
        resp = session.get(url, params={"domain": domain}, timeout=10)
        if resp.status_code != 200:
            return None, None, f"ICP lookup HTTP {resp.status_code}"
        payload = resp.json()
        if payload.get("success"):
            data = payload.get("data", {})
            return data.get("name"), data.get("icp"), "ICP lookup success"
        return None, None, payload.get("message", "ICP lookup returned no data")
    except Exception as exc:  # noqa: BLE001
        return None, None, f"ICP lookup error: {exc}"


def write_results(results: Iterable[LookupResult]) -> str:
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"ip_icp_{timestamp}.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IP ICP"
    ws.append(["IP", "Domain", "ICP Subject", "ICP License", "Message"])
    for record in results:
        ws.append([
            record.ip,
            record.domain or "",
            record.icp_subject or "",
            record.icp_license or "",
            record.message,
        ])

    wb.save(filename)
    return filename


def perform_lookup(ip_file: str, *, ipinfo_token: str | None = None, trust_env: bool = True) -> str:
    ips = read_ip_file(ip_file)
    session = requests.Session()
    session.trust_env = trust_env
    results: List[LookupResult] = []

    for ip in ips:
        domain, note = reverse_lookup(ip, session, ipinfo_token)
        if domain:
            icp_subject, icp_license, icp_note = query_icp(domain, session)
        else:
            icp_subject = icp_license = None
            icp_note = "No domain available for ICP lookup"

        message = f"{note}; {icp_note}" if note else icp_note
        results.append(LookupResult(ip, domain, icp_subject, icp_license, message))

    return write_results(results)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="IP to domain + ICP lookup")
    parser.add_argument("ip_file", help="Path to text file with one IP per line")
    parser.add_argument(
        "--no-proxy",
        action="store_true",
        help="Disable environment proxy settings for outbound requests",
    )
    parser.add_argument(
        "--ipinfo-token",
        dest="ipinfo_token",
        default=None,
        help="Optional ipinfo.io API token to raise rate limits",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output = perform_lookup(
        args.ip_file,
        ipinfo_token=args.ipinfo_token,
        trust_env=not args.no_proxy,
    )
    print(f"Results written to {output}")


if __name__ == "__main__":
    main()
