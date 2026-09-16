#!/usr/bin/env python3
"""Rerank Outlook search candidates with an OpenAI-compatible reranker gateway (vLLM).

READ-ONLY with respect to Outlook: this script never touches Outlook. It reads the JSON
produced by outlook_search.py / outlook_thread.py, sends (query, message text) pairs to the
gateway in batches, and prints the candidates sorted by relevance score.

Usage:
    python rerank.py --query "上次跟供應商談價格的信" --input candidates.json --top 10
    python rerank.py --query "..." --input candidates.json --gateway http://gw:8000/v1 --model bge-reranker-v2-m3

Configuration (CLI flag, else process environment, else the "rerank" section of .outlook-skills/settings.json
(working directory over ~/.outlook-skills), else the "env" block of ~/.claude/settings.json):
    gateway : OUTLOOK_RERANK_URL (explicit override)  >  ANTHROPIC_BASE_URL (the default: the same
              gateway Claude Code already talks to)  >  OPENAI_BASE_URL
    model   : OUTLOOK_RERANK_MODEL, default bge-reranker-v2-m3
    api key : OUTLOOK_RERANK_API_KEY  >  ANTHROPIC_AUTH_TOKEN  >  ANTHROPIC_API_KEY  >  OPENAI_API_KEY
    A more specific name wins wherever it is set. The chosen gateway is only used if the
    endpoint probe (see --endpoint auto) succeeds, so a base URL that is not a reranker
    gateway is reported as unusable instead of receiving mail data.

Endpoint formats (vLLM serves both):
    --endpoint rerank  ->  POST {gateway}/v1/rerank   {"model","query","documents":[...]}
                           response {"results":[{"index","relevance_score"}]}
    --endpoint score   ->  POST {gateway}/v1/score    {"model","text_1":query,"text_2":[...]}
                           response {"data":[{"index","score"}]}
    --endpoint auto    ->  (default) probe /v1/rerank then /v1/score with a one-document
                           request; the first one that answers correctly is used.

Privacy: subject, sender, date and body preview of every candidate are sent to the gateway.
Use only a gateway the user trusts (typically an internal vLLM instance).
"""
import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request

DEFAULT_MODEL = "bge-reranker-v2-m3"


# ---------------------------------------------------------------- config
def _settings_env():
    path = os.path.expanduser(os.environ.get("CLAUDE_SETTINGS_PATH", "~/.claude/settings.json"))
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        env = data.get("env") or {}
        return {k: str(v) for k, v in env.items()} if isinstance(env, dict) else {}
    except Exception:
        return {}


def _plugin_settings():
    """rerank section of .outlook-skills settings (local over user), mapped to the env-style names."""
    try:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import settings as ps  # settings.py next to this file
        merged, _, _, _ = ps.resolve()
        r = merged.get("rerank") or {}
        out = {}
        if r.get("gateway"):
            out["OUTLOOK_RERANK_URL"] = str(r["gateway"])
        if r.get("model"):
            out["OUTLOOK_RERANK_MODEL"] = str(r["model"])
        if r.get("api_key"):
            out["OUTLOOK_RERANK_API_KEY"] = str(r["api_key"])
        return out
    except Exception:
        return {}


def resolve_config(args):
    senv = None
    penv = _plugin_settings()

    def pick(cli, *names):
        # Specific names win over generic fallbacks regardless of where they are set:
        # OUTLOOK_RERANK_URL in settings.json beats ANTHROPIC_BASE_URL in the process env.
        nonlocal senv
        if cli:
            return cli, "cli"
        if senv is None:
            senv = _settings_env()
        for n in names:
            v = os.environ.get(n)
            if v:
                return v, f"env:{n}"
            v = penv.get(n)
            if v:
                return v, f".outlook-skills:{n}"
            v = senv.get(n)
            if v:
                return v, f"settings.json:{n}"
        return None, None

    url, url_src = pick(args.gateway, "OUTLOOK_RERANK_URL", "ANTHROPIC_BASE_URL", "OPENAI_BASE_URL")
    model, model_src = pick(args.model, "OUTLOOK_RERANK_MODEL")
    key, key_src = pick(args.api_key, "OUTLOOK_RERANK_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_API_KEY", "OPENAI_API_KEY")
    if not model:
        model, model_src = DEFAULT_MODEL, "default"
    if not url:
        raise SystemExit(
            "No gateway URL. Pass --gateway, set OUTLOOK_RERANK_URL, or add it under \"env\" in ~/.claude/settings.json."
        )
    sources = {"gateway": url_src, "model": model_src, "api_key": key_src if key else None}
    return url.rstrip("/"), model, key, sources


def endpoint_url(base: str, endpoint: str, path_override) -> str:
    if path_override:
        return base + "/" + path_override.lstrip("/")
    leaf = "rerank" if endpoint == "rerank" else "score"
    return f"{base}/{leaf}" if base.endswith("/v1") else f"{base}/v1/{leaf}"


# ---------------------------------------------------------------- input
def load_candidates(path: str):
    with open(path, "r", encoding="utf-8-sig") as fh:
        data = json.load(fh)
    if isinstance(data, list):
        return data
    for key in ("Results", "Messages", "Items"):
        if isinstance(data, dict) and isinstance(data.get(key), list):
            return data[key]
    raise SystemExit("Input JSON must be a list or an object with Results / Messages / Items.")


def doc_text(msg: dict, max_chars: int) -> str:
    subject = msg.get("Subject") or msg.get("subject") or ""
    sender = msg.get("From") or msg.get("from") or ""
    addr = msg.get("FromAddress") or ""
    date = msg.get("ReceivedTime") or msg.get("Start") or msg.get("date") or ""
    body = msg.get("BodyPreview") or msg.get("Body") or msg.get("body") or ""
    if isinstance(sender, list):  # read_msg.py format
        sender = ", ".join(a.get("name") or a.get("address") or "" for a in sender)
    head = f"Subject: {subject}\nFrom: {sender}{(' <' + addr + '>') if addr else ''}\nDate: {date}\n"
    room = max(0, max_chars - len(head))
    return head + " ".join(str(body).split())[:room]


# ---------------------------------------------------------------- gateway
def call_gateway(url, model, key, query, docs, endpoint, timeout, retries):
    if endpoint == "rerank":
        payload = {"model": model, "query": query, "documents": docs, "top_n": len(docs)}
    else:
        payload = {"model": model, "text_1": query, "text_2": docs}
    body = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if key:
        headers["Authorization"] = f"Bearer {key}"

    last = None
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, data=body, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")[:500]
            last = f"HTTP {e.code} from {url}: {detail}"
            if e.code in (400, 401, 403, 404, 422):
                raise SystemExit(last)
        except Exception as e:  # network / timeout
            last = f"{type(e).__name__}: {e}"
        if attempt < retries:
            time.sleep(2 ** attempt)
    else:
        raise SystemExit(f"Gateway call failed after {retries + 1} attempts: {last}")

    scores = [None] * len(docs)
    rows = data.get("results") if endpoint == "rerank" else data.get("data")
    if not isinstance(rows, list):
        raise SystemExit(f"Unexpected gateway response shape: {json.dumps(data)[:300]}")
    for r in rows:
        idx = r.get("index")
        sc = r.get("relevance_score", r.get("score"))
        if isinstance(idx, int) and 0 <= idx < len(docs) and sc is not None:
            scores[idx] = float(sc)
    return scores


def probe_endpoints(base, model, key, path_override, timeout):
    """Try rerank then score with a tiny request. Returns (endpoint, url, report)."""
    report = {}
    for ep in ("rerank", "score"):
        url = endpoint_url(base, ep, path_override)
        try:
            scores = call_gateway(url, model, key, "ping", ["ping"], ep, timeout, retries=0)
            if scores and scores[0] is not None:
                report[ep] = "ok"
                return ep, url, report
            report[ep] = "responded but no score in payload"
        except SystemExit as e:
            report[ep] = str(e)[:200]
    raise SystemExit(
        "No working reranker endpoint at " + base + ". "
        + "; ".join(f"{k}: {v}" for k, v in report.items())
        + ". Check OUTLOOK_RERANK_URL / OUTLOOK_RERANK_MODEL."
    )


def _select_endpoint(base, model, key, args):
    if args.endpoint == "auto":
        return probe_endpoints(base, model, key, args.path, min(args.timeout, 15.0))
    url = endpoint_url(base, args.endpoint, args.path)
    return args.endpoint, url, {args.endpoint: "forced"}


# ---------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--query", help="natural-language query (required unless --show-config)")
    ap.add_argument("--input", help="JSON from outlook_search.py (or a JSON array)")
    ap.add_argument("--gateway", help="OpenAI-compatible base URL, e.g. http://host:8000/v1")
    ap.add_argument("--model", help=f"reranker model name (default {DEFAULT_MODEL})")
    ap.add_argument("--api-key", help="bearer token if the gateway requires one")
    ap.add_argument("--endpoint", choices=["auto", "rerank", "score"], default="auto",
                    help="auto (default) probes rerank then score and uses the first that answers")
    ap.add_argument("--path", help="override the endpoint path, e.g. rerank or v2/rerank")
    ap.add_argument("--batch", type=int, default=30, help="documents per request (default 30)")
    ap.add_argument("--doc-chars", type=int, default=600, help="max characters per document (default 600)")
    ap.add_argument("--top", type=int, default=10, help="how many results to print (0 = all)")
    ap.add_argument("--min-score", type=float, default=None, help="drop results below this score")
    ap.add_argument("--timeout", type=float, default=60.0)
    ap.add_argument("--retries", type=int, default=2)
    ap.add_argument("--output", help="write JSON here instead of stdout")
    ap.add_argument("--dry-run", action="store_true", help="print the documents that would be sent, no network")
    ap.add_argument("--show-config", action="store_true", help="print the resolved gateway/model/endpoint and exit, no network")
    args = ap.parse_args(argv)

    if args.show_config:
        base, model, key, sources = resolve_config(args)
        info = {
            "base_url": base,
            "model": model,
            "api_key": "set" if key else "none",
            "sources": sources,
            "batch": args.batch,
            "doc_chars": args.doc_chars,
        }
        try:
            ep, url, report = _select_endpoint(base, model, key, args)
            info.update({"endpoint": ep, "gateway": url, "probe": report, "usable": True})
        except SystemExit as e:
            info.update({"usable": False, "error": str(e)})
        print(json.dumps(info, ensure_ascii=False, indent=2))
        sys.exit(0 if info["usable"] else 1)

    if not args.query or not args.input:
        raise SystemExit("--query and --input are required (input = JSON from outlook_search.py).")
    cands = load_candidates(args.input)
    docs = [doc_text(m, args.doc_chars) for m in cands]

    if args.dry_run:
        print(json.dumps({"count": len(docs), "documents": docs}, ensure_ascii=False, indent=2))
        return

    base, model, key, _ = resolve_config(args)
    endpoint, url, _ = _select_endpoint(base, model, key, args)

    scores = []
    batch = max(1, args.batch)
    for i in range(0, len(docs), batch):
        scores.extend(call_gateway(url, model, key, args.query, docs[i:i + batch], endpoint, args.timeout, args.retries))

    ranked = []
    for m, s in zip(cands, scores):
        if s is None:
            continue
        if args.min_score is not None and s < args.min_score:
            continue
        item = dict(m)
        item["Score"] = round(s, 4)
        ranked.append(item)
    ranked.sort(key=lambda x: x["Score"], reverse=True)
    if args.top and args.top > 0:
        ranked = ranked[: args.top]

    out = {
        "Query": args.query,
        "Gateway": url,
        "Endpoint": endpoint,
        "Model": model,
        "Candidates": len(cands),
        "Batches": (len(docs) + batch - 1) // batch,
        "Count": len(ranked),
        "Results": ranked,
    }
    text = json.dumps(out, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"Written to {args.output}")
    else:
        print(text)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
