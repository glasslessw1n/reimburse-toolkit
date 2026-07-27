#!/usr/bin/env python3
"""
Reimburse Toolkit MCP Server
用于 Hermes Agent 的 MCP (Model Context Protocol) 工具注册

在 Hermes config.yaml 中配置:
  mcp_servers:
    reimburse:
      command: python3
      args: [path/to/reimburse_mcp.py]
      description: 报销单据扫描、分类、排版工具
"""

import sys, json, os
from pathlib import Path

# 确保工具包在 Python path 中
sys.path.insert(0, str(Path(__file__).parent))

from organize import discover_trips, classify_and_sort, classify_local
from reimburse import scan_directory, scan_summary, check_didi_pairing

# MCP 工具定义
TOOLS = [
    {
        "name": "scan_bills",
        "description": "扫描报销目录，按行程汇总文件分类和数量",
        "inputSchema": {
            "type": "object",
            "properties": {
                "base_dir": {
                    "type": "string",
                    "description": "报销文件目录路径"
                }
            },
            "required": ["base_dir"]
        }
    },
    {
        "name": "check_pairing",
        "description": "核对滴滴行程单与电子发票的配对关系",
        "inputSchema": {
            "type": "object",
            "properties": {
                "base_dir": {
                    "type": "string",
                    "description": "报销文件目录路径"
                }
            },
            "required": ["base_dir"]
        }
    },
    {
        "name": "list_trips",
        "description": "列出所有行程目录和文件明细",
        "inputSchema": {
            "type": "object",
            "properties": {
                "base_dir": {
                    "type": "string",
                    "description": "报销文件目录路径"
                }
            },
            "required": ["base_dir"]
        }
    },
    {
        "name": "get_trip_detail",
        "description": "获取指定行程的详细分类数据（交通票、酒店、滴滴、餐饮）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "base_dir": {
                    "type": "string",
                    "description": "报销文件目录路径"
                },
                "trip_name": {
                    "type": "string",
                    "description": "行程目录名"
                }
            },
            "required": ["base_dir", "trip_name"]
        }
    }
]


def handle_scan_bills(args):
    base_dir = args["base_dir"]
    trips = scan_summary(base_dir)
    return {"status": "ok", "trips": str(trips)}


def handle_check_pairing(args):
    base_dir = args["base_dir"]
    ok = check_didi_pairing(base_dir)
    return {"status": "ok", "all_paired": ok}


def handle_list_trips(args):
    base_dir = args["base_dir"]
    trips, local = discover_trips(Path(base_dir))
    return {"status": "ok", "trips": trips, "local": local}


def handle_get_trip_detail(args):
    base_dir = args["base_dir"]
    trip_name = args["trip_name"]
    data = classify_and_sort(trip_name, Path(base_dir))
    if data is None:
        data = classify_local(Path(base_dir))
    if data is None:
        return {"status": "error", "message": f"行程 '{trip_name}' 未找到"}

    result = {
        "trip": trip_name,
        "transport_count": len(data.get("transport", [])),
        "hotel_pairs": len(data.get("hotel_pairs", [])),
        "didi_segments": len(data.get("didi_paired", [])),
        "dining_count": len(data.get("dining", [])),
        "telecom_count": len(data.get("telecom", [])),
    }
    return {"status": "ok", "detail": result}


HANDLERS = {
    "scan_bills": handle_scan_bills,
    "check_pairing": handle_check_pairing,
    "list_trips": handle_list_trips,
    "get_trip_detail": handle_get_trip_detail,
}


def main():
    """MCP stdio 主循环"""
    # 读取初始化请求
    for line in sys.stdin:
        try:
            request = json.loads(line.strip())
        except json.JSONDecodeError:
            continue

        method = request.get("method", "")
        req_id = request.get("id")

        if method == "initialize":
            response = {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {"tools": {}},
                    "serverInfo": {
                        "name": "reimburse-toolkit",
                        "version": "1.0.0"
                    }
                }
            }
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()

        elif method == "tools/list":
            response = {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {"tools": TOOLS}
            }
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()

        elif method == "tools/call":
            tool_name = request["params"]["name"]
            tool_args = request["params"].get("arguments", {})

            handler = HANDLERS.get(tool_name)
            if handler:
                try:
                    result = handler(tool_args)
                    content = json.dumps(result, ensure_ascii=False, default=str)
                except Exception as e:
                    content = json.dumps({"error": str(e)}, ensure_ascii=False)
            else:
                content = json.dumps({"error": f"Unknown tool: {tool_name}"})

            response = {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": content}]
                }
            }
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()

        elif method == "notifications/initialized":
            pass  # 无需响应

        else:
            response = {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32601, "message": f"Method not found: {method}"}
            }
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()


if __name__ == "__main__":
    main()
