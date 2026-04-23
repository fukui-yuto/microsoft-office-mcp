"""Microsoft Office MCP Server - PowerPoint, Word, Excel をCOMオートメーションで制御"""

from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    "Microsoft Office",
    instructions="PowerPoint、Word、ExcelをCOMオートメーションでリアルタイム制御するMCPサーバー",
)


def main():
    # Import tool modules to register @mcp.tool() decorators
    from microsoft_office.tools import powerpoint, word, excel, word_advanced, powerpoint_advanced, excel_advanced  # noqa: F401

    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
