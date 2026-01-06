import pandas as pd
import urllib.parse
import os

# ================= 编辑下面两个参数 =================
FILE_NAME = "template.xlsx"      # 你的Excel文件名
SHEET_NAME = "1"    # 你大一成绩所在的Sheet名字
# ===================================================

def generate():
    try:
        # 1. 读取班级代号 (B1单元格)
        bjdh_df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=None, nrows=1)
        bjdh = str(bjdh_df.iloc[0, 1]).strip()
        
        # 2. 从第3行开始读取表头 (header=2 表示 Excel 的第3行)
        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=2).dropna(subset=['学号'])
        
        # 大一参数修正：总指标数为 5
        mnum = len(df)
        mtymcnum = 5 
        # 对应 B2 单元格的项目 ID：1, 2, 6, 10, 11
        cate_ids = [1, 2, 6, 10, 11]
        score_cols = ['课外活动', '1000米', '运动技术1', '课内外表现', '50米']
        
        body_parts = []
        body_parts.append(f"mnum={mnum}")
        body_parts.append(f"mtymcnum={mtymcnum}")
        
        for i, (index, row) in enumerate(df.iterrows(), 1):
            # 处理学号：针对科学计数法(3.25E+09)进行修复，并补足20位空格
            try:
                # 尝试处理 3.25E+09 这种格式
                xh_raw = str(int(float(row['学号'])))
            except:
                xh_raw = str(row['学号']).split('.')[0].strip()
            
            xh_padded = xh_raw.ljust(20)
            xh_encoded = urllib.parse.quote_plus(xh_padded)
            
            # 处理姓名 (GBK)
            xm_raw = str(row['姓名']).strip()
            xm_encoded = urllib.parse.quote_plus(xm_raw.encode('gbk'))
            
            body_parts.append(f"mxh{i}={xh_encoded}")
            body_parts.append(f"mxm{i}={xm_encoded}")
            
            # 填入 5 项成绩
            for j, col in enumerate(score_cols, 1):
                val = "" if pd.isna(row[col]) else str(row[col]).strip()
                val_encoded = urllib.parse.quote_plus(val)
                body_parts.append(f"mcateid{i}{j}={cate_ids[j-1]}&mcj{i}{j}={val_encoded}")

        # 3. 拼接结尾 (注意 wtymcnum 也要同步改为 5)
        full_body = "&".join(body_parts) + f"&wnum=0&wtymcnum={mtymcnum}&Submit2=%CC%E1%BD%BB%B3%C9%BC%A8"
        
        # 4. 输出文件
        with open("大一提交内容.txt", "w") as f:
            f.write(full_body)
            
        bjdh_encoded = urllib.parse.quote(bjdh.encode('gbk'), safe='/_')
        print("✅ 大一版脚本校准完成！")
        print(f"📄 结果已存入：[大一提交内容.txt]")
        print(f"🔗 提交URL：http://tybcj.ujs.edu.cn/tea/pladdtycjdo.php?bjdh={bjdh_encoded}&teaid=please change it!")
        # please change it 改成老师的一卡通号！

    except Exception as e:
        print(f"❌ 运行失败: {e}")

if __name__ == "__main__":
    generate()
