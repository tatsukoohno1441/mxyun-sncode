import streamlit as st
import tempfile
import os
import sys
import subprocess
from pathlib import Path

# ----------- App Metadata -----------
st.set_page_config(page_title="梦想云选码工具", page_icon="📦")
st.title("梦想云选码工具")

st.markdown(
    """**使用前注意事项**

1. **库存余额表**：请先删除表格第一行（标题行）。  
2. **当日発送 CSV**：
   * 将空白 **JANコード** 补齐；
   * 删除 JAN 后面的 `-` 及其后内容；
   * 若有运费，请把运费加到单价列；
   * 完成后以 **UTF‑8** 编码保存。
    """
)

# ----------- Sidebar Inputs -----------
warehouse = st.sidebar.selectbox("选择仓库", ("通販倉庫", "なんば倉庫"))
orders_file = st.sidebar.file_uploader("上传当日発送 CSV", type=["csv"])
inv_file    = st.sidebar.file_uploader("上传库存余额表 XLSX", type=["xlsx", "xls"])

run_btn = st.sidebar.button("🚀 生成出库单")

OUTPUT_STORES = [
    "販売一丁目 Qoo10店",
    "販売一丁目 Amazon店",
    "販売一丁目 Yahoo！ショッピング店",
    "販売一丁目（楽天）",
    "ニューライフ",
    "販売一丁目 Wowma店",
]

# ----------- Processing Logic -----------

def run_make_outbound(orders_path: str, inv_path: str, warehouse: str, workdir: Path):
    """调用 make_outbound.py 生成出库单文件"""
    script = workdir / "make_outbound.py"
    if not script.exists():
        st.error("未找到 make_outbound.py， 请将脚本与 app.py 放在同一目录。")
        return
    cmd = [sys.executable, str(script), orders_path, inv_path, warehouse]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        st.error(f"脚本执行失败:\n{result.stderr}")
    else:
        st.success("生成完成！")
        st.code(result.stdout)

if run_btn:
    if not orders_file or not inv_file:
        st.warning("请先上传当日発送 CSV 和 库存余额表！")
    else:
        with st.spinner("正在处理，请稍候 …"):
            # 把上传文件保存到临时目录
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
                orders_path = tmpdir / orders_file.name
                inv_path    = tmpdir / inv_file.name
                # 保存文件
                orders_path.write_bytes(orders_file.getvalue())
                inv_path.write_bytes(inv_file.getvalue())

                # 运行脚本
                run_make_outbound(str(orders_path), str(inv_path), warehouse, Path(__file__).parent)

                # 收集输出文件并展示下载按钮
                for store in OUTPUT_STORES:
                    # 文件命名: 店铺名+行数.xlsx (未知行数，用 startswith 匹配)
                    matches = list(Path(__file__).parent.glob(f"{store}+.*.xlsx"))
                    if matches:
                        fpath = matches[0]
                        with open(fpath, "rb") as f:
                            st.download_button(
                                label=f"📥 下载 {fpath.name}",
                                data=f,
                                file_name=fpath.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                    else:
                        st.markdown(f"<span style='color:red'>【{store}】今日无订单</span>", unsafe_allow_html=True)
