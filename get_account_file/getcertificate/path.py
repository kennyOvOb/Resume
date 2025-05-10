from pathlib import Path

DESKTOP = Path.home() / 'Desktop'
BASE_DIR = DESKTOP / "資料"
STORAGE_DIR = DESKTOP / "資料-N"
GA_BASE_DIR = Path(r"Could_drive_path")
CLIENT_PORT = GA_BASE_DIR / "客戶對接"
GA_BASE_PUBLIC_DIR = GA_BASE_DIR / "06-共用資料"
GA_BASE_ACCOUNT_PUBLIC_DIR = GA_BASE_PUBLIC_DIR / "06-02-共用"  # 共用位置
DESKTOP_MOULD = DESKTOP / "模板"
CERTIFICATE_PATH = GA_BASE_DIR / "03-憑證"
DISTRIBUTE_PATH = GA_BASE_ACCOUNT_PUBLIC_DIR / "06-02-05-共用工具" / "@客戶資料.xlsx"
DATABASE_PATH = GA_BASE_ACCOUNT_PUBLIC_DIR / "06-02-05-共用工具" / "數據庫-2022.08啟用.xlsx"  # ga上的資料庫
SPOT_CHECK_PATH = GA_BASE_PUBLIC_DIR / "06-05-三組" / "抽查總表.xlsx"
OVER_BOARD_PATH = GA_BASE_ACCOUNT_PUBLIC_DIR / "06-02-02-關帳模板存放區(只會有一份)"

if __name__ == "__main__":
    pass
