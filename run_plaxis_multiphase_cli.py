import argparse
import os
import re
import sys
import time
from types import SimpleNamespace

import export_plaxis_data as core


def _now():
    return time.strftime("%H:%M:%S")


def _log(message):
    print(f"[{_now()}] {message}", flush=True)


def _error_text(exc):
    text = str(exc).strip()
    return text if text else repr(exc)


def _is_retryable(exc):
    msg = _error_text(exc).lower()
    tokens = [
        "reply code is different from what was sent",
        "request is missing",
        "code key needed for decryption",
        "decryption",
        "httpconnectionpool",
        "failed to establish a new connection",
        "winerror 10061",
        "max retries exceeded",
    ]
    return any(token in msg for token in tokens)


def _compile_regex_list(regex_list):
    out = []
    for text in regex_list:
        out.append(re.compile(text))
    return out


def _matches_any(value, compiled_regexes):
    if not compiled_regexes:
        return True
    return any(rx.search(value) for rx in compiled_regexes)


def _phase_lists(host, port, password, x_regex, y_regex):
    phases = core.list_phases_api(host, port, password)
    x_compiled = re.compile(x_regex)
    y_compiled = re.compile(y_regex)
    x_names = [p["name"] for p in phases if x_compiled.search(str(p["name"]))]
    y_names = [p["name"] for p in phases if y_compiled.search(str(p["name"]))]
    return phases, x_names, y_names


def _resolve_password(value):
    password = (value or "").strip()
    if not password:
        password = os.environ.get("PLAXIS_PASSWORD", "").strip()
    if not password:
        raise RuntimeError("Password is required. Use --password or set PLAXIS_PASSWORD.")
    return password


def _resolve_object_names(records, explicit_names, regex_list):
    by_name = {str(r.get("name", "")).strip(): r for r in records}
    resolved = []
    for name in explicit_names:
        key = str(name).strip()
        if key in by_name:
            resolved.append(key)
    if regex_list:
        compiled = _compile_regex_list(regex_list)
        for rec in records:
            name = str(rec.get("name", "")).strip()
            label = str(rec.get("label", "")).strip()
            if name and _matches_any(name, compiled):
                resolved.append(name)
                continue
            if name and label and _matches_any(label, compiled):
                resolved.append(name)
    # stable unique
    out = []
    seen = set()
    for name in resolved:
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


def _resolve_curvepoint_ids(records, explicit_ids, regex_list):
    explicit_set = set([str(v).strip() for v in explicit_ids if str(v).strip()])
    selected = []
    if explicit_set:
        for rec in records:
            rec_id = str(rec.get("id", "")).strip()
            if rec_id in explicit_set:
                selected.append(rec_id)
    if regex_list:
        compiled = _compile_regex_list(regex_list)
        for rec in records:
            rec_id = str(rec.get("id", "")).strip()
            name = str(rec.get("name", "")).strip()
            node_name = str(rec.get("node_name", "")).strip()
            label = str(rec.get("label", "")).strip()
            text = " | ".join([name, node_name, label])
            if rec_id and _matches_any(text, compiled):
                selected.append(rec_id)
    out = []
    seen = set()
    for rec_id in selected:
        if rec_id not in seen:
            seen.add(rec_id)
            out.append(rec_id)
    return out


def _run_with_retry(run_fn, attempts, sleep_sec):
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            _log(f"Run attempt {attempt}/{attempts}...")
            run_fn()
            _log("Run completed.")
            return
        except Exception as exc:
            last_exc = exc
            retryable = _is_retryable(exc)
            _log(f"Attempt {attempt} failed: {_error_text(exc)}")
            if (not retryable) or attempt >= attempts:
                raise
            _log(f"Retrying in {sleep_sec:.1f}s (retryable handshake/connection issue).")
            time.sleep(sleep_sec)
    if last_exc is not None:
        raise last_exc


def run_node_mode(args):
    password = _resolve_password(args.password)
    phases, x_names, y_names = _phase_lists(
        args.host,
        args.port,
        password,
        args.x_regex,
        args.y_regex,
    )
    _log(f"Loaded phases: total={len(phases)} x={len(x_names)} y={len(y_names)}")
    if not x_names and not y_names:
        raise RuntimeError("No phases matched x/y regex.")

    curve_records = core.list_curve_points_api(args.host, args.port, password)
    _log(f"Loaded curve points: {len(curve_records)}")
    selected_ids = _resolve_curvepoint_ids(
        curve_records,
        args.curvepoint_id,
        args.curvepoint_regex,
    )
    if selected_ids:
        _log(f"Using filtered curve points: {len(selected_ids)}")
    else:
        _log("Using all curve points.")

    run_args = SimpleNamespace(
        host=args.host,
        port=args.port,
        password=password,
        x_phase_names=x_names,
        y_phase_names=y_names,
        curvepoint_id=selected_ids,
        result_type=args.result_type,
        time_col=args.time_col,
        damping=float(args.damping),
        period_start=float(args.period_start),
        period_end=float(args.period_end),
        period_step=float(args.period_step),
        out=args.out,
    )

    _run_with_retry(
        lambda: core.run_node_multiphase_spectrum_export(run_args, logger=_log),
        attempts=args.attempts,
        sleep_sec=args.retry_sleep,
    )


def run_structural_mode(args):
    password = _resolve_password(args.password)
    phases, x_names, y_names = _phase_lists(
        args.host,
        args.port,
        password,
        args.x_regex,
        args.y_regex,
    )
    _log(f"Loaded phases: total={len(phases)} x={len(x_names)} y={len(y_names)}")
    if not x_names and not y_names:
        raise RuntimeError("No phases matched x/y regex.")

    objects = core.list_structural_objects_api(args.host, args.port, password)
    piles = objects.get("embedded_beams", [])
    plates = objects.get("plates", [])
    _log(f"Loaded structural objects: piles={len(piles)} plates={len(plates)}")

    pile_names = _resolve_object_names(piles, args.pile_name, args.pile_regex)
    p1_names = _resolve_object_names(plates, args.plate1_name, args.plate1_regex)
    p2_names = _resolve_object_names(plates, args.plate2_name, args.plate2_regex)

    if not pile_names and not p1_names and not p2_names:
        raise RuntimeError(
            "No structural object selected. Provide --pile-name/--pile-regex and/or "
            "--plate1-name/--plate1-regex and/or --plate2-name/--plate2-regex."
        )

    _log(
        f"Selected objects: piles={len(pile_names)} "
        f"plate_group1={len(p1_names)} plate_group2={len(p2_names)}"
    )

    run_args = SimpleNamespace(
        host=args.host,
        port=args.port,
        password=password,
        x_phase_names=x_names,
        y_phase_names=y_names,
        embedded_beam_names=pile_names,
        plate_group1_names=p1_names,
        plate_group2_names=p2_names,
        plate_group1_merge_single_profile=bool(args.plate1_merge_single_profile),
        plate_group2_merge_single_profile=bool(args.plate2_merge_single_profile),
        out=args.out,
    )

    _run_with_retry(
        lambda: core.run_structural_moment_export(run_args, logger=_log),
        attempts=args.attempts,
        sleep_sec=args.retry_sleep,
    )


def build_parser():
    p = argparse.ArgumentParser(
        description="Run PLAXIS multi-phase exports from terminal with retry."
    )
    sub = p.add_subparsers(dest="mode", required=True)

    def add_common(sp):
        sp.add_argument("--host", default="localhost")
        sp.add_argument("--port", type=int, default=10000)
        sp.add_argument("--password", default="", help="If empty, PLAXIS_PASSWORD env is used.")
        sp.add_argument("--x-regex", default=r"^DD2_X_.*")
        sp.add_argument("--y-regex", default=r"^DD2_Y_.*")
        sp.add_argument("--out", required=True, help="Output xlsx path.")
        sp.add_argument("--attempts", type=int, default=3)
        sp.add_argument("--retry-sleep", type=float, default=1.0)

    n = sub.add_parser("node", help="Run node multi-phase spectrum export.")
    add_common(n)
    n.add_argument("--result-type", default="Soil.Ax")
    n.add_argument("--time-col", default="DynamicTime")
    n.add_argument("--damping", type=float, default=0.05)
    n.add_argument("--period-start", type=float, default=0.01)
    n.add_argument("--period-end", type=float, default=3.0)
    n.add_argument("--period-step", type=float, default=0.01)
    n.add_argument("--curvepoint-id", action="append", default=[])
    n.add_argument(
        "--curvepoint-regex",
        action="append",
        default=[],
        help="Filter curvepoints by regex on name/node/label. Can repeat.",
    )

    s = sub.add_parser("structural", help="Run structural moment export.")
    add_common(s)
    s.add_argument("--pile-name", action="append", default=[])
    s.add_argument("--pile-regex", action="append", default=[])
    s.add_argument("--plate1-name", action="append", default=[])
    s.add_argument("--plate1-regex", action="append", default=[])
    s.add_argument(
        "--plate1-merge-single-profile",
        action="store_true",
        help="Treat selected Plate Group 1 elements as a single merged profile.",
    )
    s.add_argument("--plate2-name", action="append", default=[])
    s.add_argument("--plate2-regex", action="append", default=[])
    s.add_argument(
        "--plate2-merge-single-profile",
        action="store_true",
        help="Treat selected Plate Group 2 elements as a single merged profile.",
    )
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()
    try:
        if args.mode == "node":
            run_node_mode(args)
        elif args.mode == "structural":
            run_structural_mode(args)
        else:
            raise RuntimeError(f"Unknown mode: {args.mode}")
    except KeyboardInterrupt:
        _log("Interrupted by user.")
        sys.exit(130)
    except Exception as exc:
        _log(f"FAILED: {_error_text(exc)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
