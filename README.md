# 🧪 Test Case Manager

Python, **CustomTkinter** ve **SQLite** ile geliştirilmiş, test senaryolarınızı tek bir masaüstü uygulamasından oluşturmanızı, düzenlemenizi ve **Excel'e** aktarmanızı sağlayan bir test case yönetim aracı.

---

## 📌 Özellikler

- 📝 **Test Senaryosu Oluşturma / Düzenleme** — TC ID, isim, öncelik, uygulama, versiyon, oluşturan kişi ve tarih gibi meta bilgilerle birlikte
- ✅ **Precondition / Postcondition** listeleri — dinamik satır ekleme/silme
- 📊 **Test Data** yönetimi — key/value çiftleri halinde
- 🔢 **Test Steps** — adım no, aksiyon, test verisi ve beklenen sonuç
- 🤖 **Otomasyon durumu** işaretleme (Automated / Manual)
- 📤 **Excel'e aktarım** — her test senaryosu için ayrı, biçimlendirilmiş bir detay sayfası ve hyperlink'li bir özet sayfası içeren `.xlsx` raporu
- 🗑️ Test senaryosu **silme** (bağlı tüm precondition/postcondition/data/step kayıtlarıyla birlikte)
- 💾 Yerel **SQLite** veritabanı — kuruluma gerek yok, harici sunucu bağımlılığı yok

---

## 🖥️ Ekran Görünümü

Uygulama koyu temalı (dark mode) bir arayüz kullanır: solda test listesi ve arama alanı, sağda seçilen testin tüm detaylarının göründüğü kaydırılabilir bir panel bulunur.

---

## 🛠️ Kullanılan Teknolojiler

| Katman | Teknoloji |
|---|---|
| Arayüz (GUI) | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| Veritabanı | SQLite3 (Python standart kütüphanesi) |
| Excel Raporlama | [openpyxl](https://openpyxl.readthedocs.io/) |
| Dil | Python 3.10+ |

---

## 📁 Proje Yapısı

```
TestCaseManager/
├── main.py                  # Uygulama giriş noktası
├── requirements.txt         # Python bağımlılıkları
├── data/
│   └── testcases.db         # SQLite veritabanı (otomatik oluşturulur, git'e dahil değildir)
├── database/
│   ├── database.py          # SQLite bağlantısı ve tüm CRUD sorguları
│   └── models.py            # TestCase / TestStep / TestData veri modelleri (dataclass)
├── excel/
│   ├── exporter.py          # Excel export mantığı (openpyxl)
│   └── template.xlsx        # Export için kullanılan şablon dosyası
└── gui/
    ├── app.py                # Ana pencere (CTk)
    ├── dashboard.py           # Sol panel (liste) + sağ panel (detay) ekranı
    └── new_test.py            # Yeni test oluşturma / düzenleme penceresi
```

---

## 🚀 Kurulum

### 1. Depoyu klonlayın

```bash
git clone https://github.com/sevvalcrt/Test-Case-Manager.git
cd Test-Case-Manager
```

### 2. Sanal ortam oluşturun (önerilir)

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. Bağımlılıkları yükleyin

```bash
pip install -r requirements.txt
```

### 4. Uygulamayı çalıştırın

```bash
python main.py
```

> İlk çalıştırmada `data/` klasörü altında `testcases.db` adında boş bir SQLite veritabanı otomatik olarak oluşturulur.

---

## 📖 Kullanım

1. Sol panelde **"➕ Yeni Test"** butonuna tıklayarak yeni bir test senaryosu formu açın.
2. TC ID ve Test Adı zorunlu alanlardır; geri kalan bilgileri (precondition, test data, adımlar, postcondition) doldurun.
3. Kaydettikten sonra test, sol panelde listelenir; üzerine tıklayarak sağ panelde tüm detaylarını görüntüleyebilirsiniz.
4. **"✏ Düzenle"** ile mevcut bir testi güncelleyebilir, **"🗑 Sil"** ile silebilirsiniz.
5. **"📤 Excel'e Aktar"** butonuyla tüm test senaryolarını, her biri için ayrı bir detay sayfası içeren `export.xlsx` dosyasına aktarabilirsiniz.

---

## 🗄️ Veritabanı Şeması

| Tablo | Açıklama |
|---|---|
| `test_cases` | Ana test senaryosu bilgileri (tc_id, name, priority, application, version, creator, create_date, automation) |
| `pre_conditions` | Her test senaryosuna bağlı ön koşullar |
| `test_data` | Key/value formatında test verileri |
| `test_steps` | Sıralı test adımları (step_no, action, test_data, expected_result) |
| `post_conditions` | Her test senaryosuna bağlı son koşullar |

Tüm alt tablolar `tc_id` üzerinden `test_cases` tablosuna foreign key ile bağlıdır.

---

## 🤝 Katkıda Bulunma

Katkılarınızı memnuniyetle karşılarız:

1. Bu repoyu fork'layın
2. Yeni bir branch oluşturun (`git checkout -b ozellik/yeni-ozellik`)
3. Değişikliklerinizi commit'leyin (`git commit -m 'Yeni özellik eklendi'`)
4. Branch'inizi push'layın (`git push origin ozellik/yeni-ozellik`)
5. Bir Pull Request açın

---

## 📄 Lisans

Bu proje için henüz bir lisans belirlenmemiştir. Eklemek isterseniz [choosealicense.com](https://choosealicense.com/) üzerinden projenize uygun bir lisans (örn. MIT) seçebilirsiniz.

---

## 👤 Geliştirici

**sevvalcrt** — [GitHub](https://github.com/sevvalcrt)