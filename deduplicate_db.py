import sqlite3
import os

DB_PATH = r'c:\Users\selcu\OneDrive\Masaüstü\uretim_takip\uretim.db'

def deduplicate():
    if not os.path.exists(DB_PATH):
        print("Database not found.")
        return

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # 1. Mükerrer olan (boşluksuz halleri aynı olan) grupları bul
    c.execute("""
        SELECT REPLACE(referans_kodu, ' ', ''), COUNT(*) 
        FROM referans_listesi 
        GROUP BY REPLACE(referans_kodu, ' ', '') 
        HAVING COUNT(*) > 1
    """)
    duplicates = c.fetchall()
    
    print(f"Tespit edilen mükerrer grup sayısı: {len(duplicates)}")

    for norm_code, count in duplicates:
        print(f"\nTemizleniyor: {norm_code} ({count} adet)")
        
        # Bu gruptaki tüm kayıtları getir
        c.execute("SELECT id, referans_kodu, hedef_cycle_time_sn FROM referans_listesi WHERE REPLACE(referans_kodu, ' ', '') = ?", (norm_code,))
        rows = c.fetchall()
        
        # En 'düzgün' olanı seç (Süresi olanı veya en uzun isimlisi genellikle daha açıklayıcıdır)
        # Öncelik: hedef_cycle_time_sn > 0 olan, yoksa ilk kayıt
        main_row = None
        for row in rows:
            if row[2] > 0:
                main_row = row
                break
        if not main_row:
            main_row = rows[0]
            
        print(f"  Tutulacak Kayıt: {main_row[1]} (id: {main_row[0]}, süre: {main_row[2]})")
        
        # Diğerlerini sil
        for row in rows:
            if row[0] != main_row[0]:
                print(f"  Siliniyor: {row[1]} (id: {row[0]})")
                c.execute("DELETE FROM referans_listesi WHERE id = ?", (row[0],))
                
    conn.commit()
    conn.close()
    print("\nVeritabanı temizliği tamamlandı.")

if __name__ == "__main__":
    deduplicate()
