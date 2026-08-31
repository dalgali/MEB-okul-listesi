# MEB Okul Listesi

Türkiye'deki **bütün resmî okulların ve MEB kurumlarının** listesi.
Millî Eğitim Bakanlığı'nın kendi
[okullar sayfasından](https://www.meb.gov.tr/baglantilar/okullar/index.php)
otomatik olarak çekilir.

**İçinde ne var:** 55.216 okul/kurum — her biri için il, ilçe, ad, kurum
kodu, okul türü, internet sitesi ve harita bağlantısı.
Son güncelleme: **31 Ağustos 2026**.

---

## Sadece veri lazımsa (kod bilmene gerek yok)

`data/` klasörüne gir, sana uygun dosyayı indir. Dördü de **aynı bilgiyi**
içerir, sadece format farklı:

| Dosya | Ne zaman kullan |
|---|---|
| **`data/meb-okullar.csv`** | Excel veya Google Sheets'te açacaksan. En yaygın seçenek. |
| **`data/meb-okullar.xlsx`** | Doğrudan Excel dosyası istiyorsan (çift tıkla, açılır). |
| **`data/meb-okullar.json`** | Bir programa / koda veri vereceksen. |
| **`data/meb-okullar.sqlite`** | Veritabanı sorgusu (SQL) yazacaksan. |

> CSV'yi Excel'de açınca Türkçe harfler bozuk görünürse: dosyayı Excel'de
> "Veri > Metinden/CSV'den" ile içe aktar, kodlama olarak **UTF-8** seç.
> (Dosya zaten UTF-8, bu sadece Excel'in huysuzluğu.)

### Tablodaki sütunlar ne anlama geliyor

| Sütun | Örnek | Açıklama |
|---|---|---|
| `il` | `Ankara` | Okulun ili. (Kaynak `ANKARA` diye verir; okunaklı yazıma çevrildi.) |
| `ilce` | `Akyurt` | Okulun ilçesi. |
| `okul_adi` | `Akyurt Anadolu Lisesi` | Kurumun tam adı. |
| `kurum_kodu` | `966556` | MEB'in her okula verdiği 6 haneli numara. Her okulda dolu ve benzersiz. |
| `okul_turu` | `Anadolu Lisesi` | Okulun türü. **Tahmindir** — adına bakılarak bulunur (aşağıda açıklandı). |
| `web_sitesi` | `https://akyurtanadolu.meb.k12.tr` | Okulun resmî sitesi. |
| `adres` | *(boş)* | Bu listede yok. Nedeni aşağıda. |
| `telefon` | *(boş)* | Bu listede yok. Nedeni aşağıda. |
| `harita` | `https://akyurtanadolu.meb.k12.tr/tema/harita.php` | Okulun konum sayfası. Linke tıklayınca haritada açılır. |
| `cekim_tarihi` | `2026-08-31` | Verinin MEB'den çekildiği gün. |

### "adres" ve "telefon" neden boş?

MEB'in okul listesi sayfası bu ikisini **vermiyor**. Adres ve telefon
her okulun kendi ayrı sitesinde (`...meb.k12.tr/tema/iletisim.php`) yazıyor.
55 binden fazla siteyi tek tek gezmek gerekir; bu ayrı ve uzun bir iş.
Sütunlar ileride doldurulabilsin diye tabloda **boş bırakıldı**, silinmedi.

### "okul_turu" nasıl bulunuyor?

MEB listede okul türünü ayrı bir bilgi olarak vermiyor. Bu yüzden okulun
**adına bakıyoruz**: adında "Anadolu Lisesi" geçiyorsa → `Anadolu Lisesi`,
"İlkokulu" geçiyorsa → `İlkokul`, "Halk Eğitimi Merkezi" geçiyorsa →
`Halk Eğitimi Merkezi`... Bu bir **tahmin**, resmî sınıflandırma değil.
55.216 kurumun 55.203'ü otomatik sınıflandı; adı hiçbir kalıba uymayan
13 idari birim `Belirsiz` olarak işaretli.

### `il` sütununda `Bakanlık` yazan satırlar

Bir ile bağlı olmayan kayıtlar: Bakanlık merkez birimleri, yurt dışındaki
Türk okulları, döner sermaye işletmeleri gibi. Sadece "gerçek" okulları
istiyorsan bu satırları ele: `il` sütunu `Bakanlık` olanları çıkar.

### Listede kaç tane, ne var? (okul türü dağılımı)

```
22.747  İlkokul                            381  Fen Lisesi
13.025  Ortaokul                           372  Bilim ve Sanat Merkezi
 4.155  Anaokulu                           303  Rehberlik ve Araştırma Merkezi
 2.912  Anadolu Lisesi                     109  Lise (genel)
 2.520  İmam Hatip Ortaokulu               107  Spor Lisesi
 2.461  Mesleki ve Teknik Anadolu Lisesi   105  Güzel Sanatlar Lisesi
 1.755  Anadolu İmam Hatip Lisesi           97  Sosyal Bilimler Lisesi
 1.496  Özel Eğitim Kurumu                  32  Olgunlaşma Enstitüsü
 1.002  Halk Eğitimi Merkezi                13  Belirsiz
   583  Çok Programlı Anadolu Lisesi       ...  (küçük idari birimler)
   536  Öğretmenevi
   475  Mesleki Eğitim Merkezi
```

---

## Listeyi kendin güncellemek istersen (kod çalıştırma)

Veri eskidiyse veya baştan çekmek istersen `meb_okullar.py` scriptini
çalıştırırsın. Adım adım:

### 1. Python kurulu mu bak

Terminal (Mac'te "Terminal", Windows'ta "Komut İstemi") aç, şunu yaz:

```bash
python3 --version
```

`Python 3.9` veya üstü yazıyorsa tamam. Yazmıyorsa
[python.org](https://www.python.org/downloads/)'dan kur.

### 2. Bu projeyi indir

```bash
git clone https://github.com/dalgali/MEB-okul-listesi.git
cd MEB-okul-listesi
```

(Git yoksa: GitHub'da yeşil **"Code"** düğmesi → **"Download ZIP"**, sonra
zip'i aç ve o klasöre gir.)

### 3. Gerekli paketleri kur

```bash
pip install -r requirements.txt
```

Bu iki şey kurar: `requests` (internetten veri çeker), `openpyxl` (Excel
dosyası yazar).

### 4. Scripti çalıştır

```bash
python3 meb_okullar.py --out-dir data
```

Ekranda ilerleme çubuğu akar (`1000/55216`, `2000/55216` ...). Bir
dakikadan kısa sürer. Bitince `data/` klasöründeki 4 dosya yenilenmiş olur.

### İşin bittiğinde göreceğin çıktı (örnek)

```
MEB'e bağlanılıyor…
Okullar çekiliyor — tüm Türkiye
    1000/55216 (% 1.8)
    ...
   55216/55216 (%100.0)

Toplam benzersiz kayıt: 55216  (yinelenen atlandı: 0)
  kurum_kodu boş: 0  |  web_sitesi boş: 0  |  türü 'Belirsiz': 13
  yazıldı: data/meb-okullar.csv  (9.107 KB)
  yazıldı: data/meb-okullar.json  (15.848 KB)
  yazıldı: data/meb-okullar.xlsx  (3.185 KB)
  yazıldı: data/meb-okullar.sqlite  (11.392 KB)
Tamamlandı.
```

### Kullanabileceğin ek ayarlar

| Yazarsan | Ne yapar |
|---|---|
| `python3 meb_okullar.py --il 81` | Sadece tek il çeker (81 = Düzce, ~300 okul). Hızlı deneme için. |
| `python3 meb_okullar.py --formats csv` | Sadece CSV üretir (json/xlsx/sqlite istemiyorsan). |
| `python3 meb_okullar.py --out-dir masaustu` | Dosyaları başka klasöre yazar. |
| `python3 meb_okullar.py --selftest` | İnternete çıkmadan scriptin sağlamlığını test eder. |
| `python3 meb_okullar.py --help` | Bütün ayarları listeler. |

---

## Sık sorulanlar

**Bu yasal mı?** Veri MEB'in herkese açık sayfasından geliyor, giriş/şifre
yok. Script kibar davranır: sayfalar arasında bekler, tek seferde ~56
istek atar.

**Script çalışmıyor / "Erişim yetkiniz yok" diyor.** MEB sitesi geçici
olarak kapalı veya yapısını değiştirmiş olabilir. Birkaç saat sonra
tekrar dene; sorun sürerse Issues bölümünden bildir.

**Veri ne kadar günceldir?** `data/` içindeki dosyalar `cekim_tarihi`
sütunundaki tarihte çekildi. Daha güncel istiyorsan yukarıdaki adımlarla
kendin çalıştır.

**İlçe adları neden hep aynı yazımda?** Kaynakta bazı kayıtlar ilçeyi
Türkçe karakter içermeyen biçimde giriyor (`Bağcılar` yerine `Bagcılar`,
`Kadıköy` yerine `Kadiköy`). Script bunları aynı ildeki en sık geçen doğru
yazıma eşitler; böylece İstanbul 44 değil 39 ilçe gösterir.

**Teknik detay:** MEB sayfası listeyi `okullar_ajax.php` adresinden JSON
olarak alıyor; `Referer` başlığı zorunlu. Script bu adresi 1000'erlik
sayfalarla gezer. `il` / `ilce` / `okul_adi`, gelen `OKUL_ADI` metninin
`"İL - İLÇE - Ad"` kalıbından ayrıştırılır (il/ilçe `ANKARA` → `Ankara`
biçimine, Türkçe `İ`/`I` kurallarına uyularak çevrilir); `kurum_kodu`,
`YOL` alanının son parçasından (`06/26/966556`) alınır.

---

## Lisans

MIT — istediğin gibi kullan, değiştir, dağıt. Bkz. [LICENSE](LICENSE).
