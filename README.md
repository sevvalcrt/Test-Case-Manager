# Test Case Manager

Python ve CustomTkinter kullanılarak geliştirilmiş, **manuel test senaryolarını oluşturma, düzenleme, yönetme ve Excel formatında dışa aktarma** amacı taşıyan masaüstü Test Case Management uygulamasıdır.

## Özellikler

* 📁 Proje oluşturma ve proje yönetimi
* 🧪 Test case oluşturma, düzenleme ve silme
* 🔎 Test case arama
* 🏷️ Test durumu yönetimi
* 🖥️ Test türü ve test ortamı bilgileri
* 📝 Pre-condition, test data, test steps ve post-condition yönetimi
* 🤖 Otomasyon durumu ve otomasyon senaryosu takibi
* 🐞 Hata kodu ve hata önceliği bilgileri
* 📊 Test case'leri Excel formatında dışa aktarma
* 🔗 Excel içerisinde test case'ler arasında hyperlink desteği
* 💾 Her proje için ayrı SQLite veritabanı
* 🖥️ Modern CustomTkinter arayüzü

## Kullanılan Teknolojiler

* **Python**
* **CustomTkinter**
* **SQLite**
* **OpenPyXL**
* **Tkinter**

## Proje Yapısı

```text
TestCaseManager/
│
├── database/
│   ├── database.py
│   └── models.py
│
├── excel/
│   └── exporter.py
│
├── gui/
│   ├── app.py
│   ├── dashboard.py
│   ├── new_test.py
│   └── project_manager.py
│
├── projects/
│   └── <ProjectName>/
│       └── testcases.db
│
├── main.py
├── requirements.txt
└── README.md
```

## Kurulum

Projeyi bilgisayarınıza klonlayın:

```bash
git clone <repository-url>
cd TestCaseManager
```

Gerekli Python paketlerini yükleyin:

```bash
pip install -r requirements.txt
```

## Uygulamayı Çalıştırma

Projeyi çalıştırmak için:

```bash
python main.py
```

Uygulama açıldığında proje yönetim ekranı üzerinden mevcut bir projeyi seçebilir veya yeni bir proje oluşturabilirsiniz.

## Proje Yönetimi

Her proje kendi SQLite veritabanına sahip olacak şekilde saklanır.

Örneğin:

```text
projects/
├── Mango/
│   └── testcases.db
│
└── ExampleProject/
    └── testcases.db
```

Bu yapı sayesinde farklı projelerin test case'leri birbirinden bağımsız olarak yönetilebilir.

## Test Case Bilgileri

Her test case aşağıdaki bilgileri içerebilir:

* Test Case ID
* Test Case Name
* Test Type
* Test Environment
* Status
* Pre-Conditions
* Test Data
* Test Steps
* Post-Conditions
* Automation Requested
* Automation Completed
* Automation Scenario
* Error Code
* Error Priority

## Test Case Durumları

Test case'ler aşağıdaki durumlardan biriyle takip edilebilir:

* **Not Run**
* **Passed**
* **Failed**
* **Blocked**

## Excel Export

Test case'ler Excel formatında dışa aktarılabilir.

Export işlemi sonucunda:

* Test case özetlerinin bulunduğu bir ana sayfa
* Her test case için ayrı detay sayfası
* Test case'ler arasında hyperlinkler
* Test bilgileri
* Pre-condition
* Test data
* Test steps
* Post-condition
* Automation bilgileri
* Hata bilgileri

oluşturulur.

Excel dosyası manuel test sonuçlarının raporlanması ve paylaşılması için kullanılabilir.

## Requirements

Projenin çalışması için gerekli temel paketler:

```text
customtkinter>=5.2.2
openpyxl>=3.1.5
```

Python'un güncel bir sürümünün kullanılması önerilir.

## Kullanım

1. Uygulamayı başlatın.
2. Bir proje seçin veya yeni proje oluşturun.
3. Dashboard üzerinden test case'leri görüntüleyin.
4. Yeni test case oluşturun veya mevcut bir test case'i düzenleyin.
5. Test adımlarını ve gerekli test bilgilerini ekleyin.
6. Test durumunu güncelleyin.
7. Gerektiğinde test case'leri Excel formatında dışa aktarın.

## Amaç

Bu proje, manuel test süreçlerinin daha düzenli şekilde yönetilmesini sağlamak ve test case'lerin merkezi bir masaüstü uygulaması üzerinden takip edilmesine yardımcı olmak amacıyla geliştirilmiştir.

Ayrıca ilerleyen aşamalarda manuel test süreçlerinin otomasyon ile desteklenebilmesi için **automation status** ve **automation scenario** gibi alanlar da sisteme dahil edilmiştir.

## Geliştirici

**Şevval**

Software Engineering Student
