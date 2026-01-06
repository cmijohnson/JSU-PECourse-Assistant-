import pandas as pd
import urllib.parse
import os

# ================= 编辑下面两个参数 =================
FILE_NAME = "template.xlsx"      # 你的Excel文件名
SHEET_NAME = "1"    # 你要处理的那个Sheet名字
# ===================================================

def generate():
    try:
        # 1. 读取班级代号 (B1单元格)
        bjdh_df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=None, nrows=1)
        bjdh = str(bjdh_df.iloc[0, 1]).strip()
        
        # 2. 从第4行开始读取数据
        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=2).dropna(subset=['学号'])
        
        # 系统固定参数
        mnum = len(df)
        mtymcnum = 6
        cate_ids = [1, 2, 6, 8, 10, 11]
        score_cols = ['课外活动', '1000米', '运动技术1', '立定跳远', '课内外表现', '50米']
        # 此处可以根据实际情况进行修改！
        
        body_parts = []
        body_parts.append(f"mnum={mnum}")
        body_parts.append(f"mtymcnum={mtymcnum}")
        
        for i, (index, row) in enumerate(df.iterrows(), 1):
            # --- 修正点 1: 使用 quote_plus 确保空格转为 + 号，且固定20位长度 ---
            xh_raw = str(row['学号']).split('.')[0].strip()
            xh_padded = xh_raw.ljust(20) # 填充空格至20位
            xh_encoded = urllib.parse.quote_plus(xh_padded)
            
            # --- 修正点 2: 姓名 GBK 编码 ---
            xm_raw = str(row['姓名']).strip()
            xm_encoded = urllib.parse.quote_plus(xm_raw.encode('gbk'))
            
            body_parts.append(f"mxh{i}={xh_encoded}")
            body_parts.append(f"mxm{i}={xm_encoded}")
            
            # 填入6项成绩
            for j, col in enumerate(score_cols, 1):
                val = "" if pd.isna(row[col]) else str(row[col]).strip()
                # 去掉成绩里的空格
                val_encoded = urllib.parse.quote_plus(val)
                body_parts.append(f"mcateid{i}{j}={cate_ids[j-1]}&mcj{i}{j}={val_encoded}")

        # 拼接结尾
        full_body = "&".join(body_parts) + "&wnum=0&wtymcnum=6&Submit2=%CC%E1%BD%BB%B3%C9%BC%A8"
        
        # 3. 输出文件
        with open("POST 提交内容.txt", "w") as f:
            f.write(full_body)
            
        # 4. 生成 URL (注意 bjdh 的编码方式)
        # 系统样本中 bjdh 里的 / 没有被编码，所以使用 safe='/'
        bjdh_encoded = urllib.parse.quote(bjdh.encode('gbk'), safe='/_')
        final_url = f"http://tybcj.ujs.edu.cn/tea/pladdtycjdo.php?bjdh={bjdh_encoded}&teaid=please change it"
        # please change it 改成老师的一卡通号！
        
        print("✅ 深度校准完成！")
        print(f"📄 1. 请打开 [提交内容.txt]，全选复制里面的内容。")
        print(f"🔗 2. 抓包工具中的目标 URL 应为：\n{final_url}")

    except Exception as e:
        print(f"❌ 运行失败: {e}")

if __name__ == "__main__":
    generate()
