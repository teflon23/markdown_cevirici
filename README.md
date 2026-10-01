# Yerel PDF → Markdown ve doğrulama hattı

Tamamen yerel çalışır. Ağ istemcisi, bulut API'si, model indirme veya LLM ile metin düzeltme içermez. PDF içindeki metinler veri olarak işlenir; talimat olarak yürütülmez. Qwen 8B gerektirmez; çıktıyı daha sonra yerel RAG sisteminize verebilirsiniz.

**Önemli sınır:** PDF'nin görsel anlamının yüzde yüz korunduğu otomatik olarak kanıtlanamaz. Bu araç kaynak metin katmanındaki kelimelerin muhasebesini kanıtlar, geometrik okumayı denetler ve insan incelemesi olmadan `ready.md` oluşturmaz. Kelime sayılarının eşit olması, tablo ilişkilerinin veya okuma sırasının doğru olduğunu tek başına göstermez.

## Çalıştırma

Python 3.10+ gerekir. `run.ps1` sırasıyla proje içindeki `.venv`, PATH üzerindeki Python ve Windows `py` başlatıcısını kullanır. Yapay zekâ eklentisi gerekmez. Bir üretim ortamında ayrılmış bir sanal ortam kullanın.

```powershell
.\run.ps1 convert "belge.pdf" ".\output\yeni-calistirma"
.\run.ps1 verify ".\output\yeni-calistirma"
```

Standart Python ortamında:

```text
python pdf_to_markdown.py convert source.pdf output/run-001
python pdf_to_markdown.py verify output/run-001
python -m unittest -v test_pipeline
```

Çıktı dizini yeni olmalıdır; mevcut çalışma üzerine yazılmaz. Çıkış kodları: `0` doğrulama/yayımlama başarılı, `1` çalışma hatası, `2` kelime muhasebesi başarısız, `3` dönüşüm tamamlandı fakat anlamsal inceleme gerekli. `verify` komutunun `0` dönmesi, anlamsal doğruluk onayı değildir. Marj bandı gerekirse `--margin-fraction 0.12` ile daraltılabilir; varsayılan üst/alt yüzde 16'dır. Sayfa görsellerinin çözünürlüğü `--dpi 150` ile değiştirilebilir.

## Artifactory üzerinden kurulum

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --index-url "https://ARTIFACTORY-SUNUCUSU/artifactory/api/pypi/PYTHON-DEPOSU/simple" -r requirements.txt
.\.venv\Scripts\python.exe -m unittest -v test_pipeline
```

Adres yer tutucudur; kurumunuzun verdiği pip index adresiyle değiştirin. Kurumunuz pip'i zaten Artifactory için yapılandırdıysa `--index-url` parametresini atlayabilirsiniz. Ek genel paket deposu tanımlamayın; kimlik doğrulama ve kurum sertifikaları için kurumunuzun yöntemini kullanın. Şifre, token ve yerel pip yapılandırmasını Git'e eklemeyin.

## Air-gap kurulumu

Yeni air-gap makine için hedef işletim sistemi ve Python sürümüyle uyumlu bağımlılık paketlerini güvenilir hazırlık makinesinde kurumunuzun onayladığı depodan önceden temin edin. Hazırlık aşaması ağ erişimi gerektirebilir; air-gap çalıştırma aşaması gerektirmez.

```text
# Yalnızca dışarıdaki hazırlık makinesinde:
python -m pip download --index-url "https://ARTIFACTORY-SUNUCUSU/artifactory/api/pypi/PYTHON-DEPOSU/simple" --dest wheelhouse -r requirements.txt
# Wheelhouse ile requirements.txt dosyasını kontrollü şekilde hedefe taşıdıktan sonra:
python -m pip install --no-index --find-links=wheelhouse -r requirements.txt
```

Doğrudan bağımlılıklar sabit sürümlüdür. Üretim dağıtımı için transitive wheel dosyalarını da sürüm/hash manifestiyle sabitleyin; sistem/paket tedarik güvenliği bu betiğin kapsamından ayrıdır. PDF render işlemi pdfplumber'ın yerel PDFium bağımlılığıyla yapılır; tarayıcı veya harici servis kullanılmaz.

## Aşamalar ve çıktılar

1. Kaynak PDF kopyalanır ve SHA-256 alınır. İlk geçişte yalnızca üst/alt marjda tekrar eden satırlar öğrenilir.
2. İkinci geçişte her sayfanın karakterleri, kelimeleri, koordinatları ve dışlanan marj kelimeleri kaydedilir.
3. Tablo kenarları aranır. Word kaynaklı PDF'lerde ince dolu dikdörtgenler kenar kabul edilir; geniş hücre arka planları sahte kenar sayılmaz. Kelimeler merkez koordinatına göre hücrelere atanır. Çakışmalı/sahipsiz hücrelerde tablo metin olarak korunur ve sorun işaretlenir.
4. Metin satırları birleştirilir; paragraf boşlukları, hukuki madde işaretleri ve tablo sınırları korunur. `CHAPTER`, `ANNEX`, `Article 25.2`, `(1)`, `(g)`, `(iii)` hiyerarşisi Markdown başlıklarına dönüştürülür. Her bloğun bağlamı sayfalar arasında taşınır.
5. Ayrı bir yerel ayrıştırıcı olan pypdf ile karakter envanteri karşılaştırılır. Karşılaştırmada NFKC/boşluk normalizasyonu kullanılır; Markdown içeriğine NFKC uygulanmaz. Bu karşılaştırma sıra doğrulaması değildir ve iki ayrıştırıcı aynı hatayı yapabilir.
6. Her kaynak kelime, Markdown'a tam bir kez yazılmış veya gerekçeli dışlama kaydında bulunmuş olmalıdır. Markdown'dan geri çözülen içerik, blok kelimeleriyle **sıralı** karşılaştırılır. Son dosya bayt aralıkları ve hash'leri yeniden doğrulanabilir.
7. Her sayfa yerel PNG olarak render edilir. Anlamsal kabul ayrıca incelenir.

| Dosya | Kullanımı |
|---|---|
| `source.pdf` | Değiştirilmemiş kaynak; görsellerin ve bütün özgün bilginin arşivi |
| `candidate.md` | İnceleme bekleyen Markdown; otomatik ingestion için hazır sayılmaz |
| `evidence.jsonl.gz` | Sayfa başına karakter, kelime, bbox, tablo hücresi, hiyerarşi, dışlama, Markdown bayt aralığı ve pypdf metni |
| `report.json` | Kaynak/çıktı hash'leri, sürümler, ayarlar, sayfa sorunları ve başarısızlıklar |
| `review/page-0001.png` vb. | Görsel karşılaştırma için kaynak sayfalar |
| `review/index.html` | `build_review.py` ile oluşturulan, kaynak görseli ve çıkarılan içeriği yan yana sunan çevrimdışı kontrol ekranı |
| `review-template.json` | Kaynak/çıktı hash'lerine bağlı sayfa inceleme şablonu |
| `ready.md` | Yalnızca açık yerel inceleme onayından sonra oluşturulur |
| `RUNNING.json` | Varsa çalışma tamamlanmamıştır; ingest etmeyin |

Tablo sütunları için varsayılan olarak `Column 1` gibi **üretilmiş** başlıklar kullanılır; kaynak ilk satır aynen veri satırı olarak tutulur. Böylece veri satırını yanlışlıkla başlık saymayız. Bu başlıklar kelime muhasebesinden ayrıdır. Birleşik hücreler Markdown'da tam ifade edilemediğinden, kaynak geometrisi sidecar kaydında korunur ve inceleme istenir. Hücre metni tekrar edilmez; boş hücreyi doldurmak için tahmin yapılmaz. Sayfa sınırında tablolar otomatik birleştirilmez.

## Kontrol ve kabul

`report.json` içindeki `failures` gerçek kaynak/çıktı muhasebe hatalarıdır; `issues` görsel/anlamsal belirsizliklerdir. Metin katmanı olmayan sayfalar, eksik glifler, grafikler, çizgisiz tablo/çoklu sütun olasılığı, birleşik hücreler, tire ve sayfa devamları incelemeye işaretlenir. İşaret bulunmaması semantik garanti değildir; tüm sayfalar onay gerektirir.

Yan yana kontrol ekranını üretmek için `python build_review.py output/run-001` çalıştırın ve `review/index.html` dosyasını yerel tarayıcıda açın. Bu ekran JavaScript, sunucu veya ağ kullanmaz; kaynak metni HTML olarak çalıştırmaz.

`review-template.json` dosyasının bir kopyasında `reviewer` alanını doldurun. Her sayfayı kaynak görselle karşılaştırdıktan sonra `approved: true` ve açıklayıcı `note` girin. Toplu olarak görmeden onaylamayın.

```text
python pdf_to_markdown.py release output/run-001 --review output/run-001/my-review.json
```

Kaynak/Markdown/evidence değişmişse, eksik sayfa varsa veya muhasebe başarısızsa kabul engellenir. İnceleme bir insan beyanıdır; kriptografik imza veya dışarıdan değiştirilemez denetim sistemi değildir. Dosyaları değiştirip yeniden hash üretmek anlamsal bir düzeltme süreci oluşturmaz. Düzeltmeler için kaynak/ayarlar üzerinden yeni çalışma üretin ve tekrar inceleyin.

## RAG kullanımı

Yalnızca kabul edilen `ready.md` dosyasını indeksleyin. MarkdownHeaderTextSplitter için önerilen başlıklar:

```python
headers_to_split_on = [
    ("#", "chapter"), ("##", "article_or_section"),
    ("###", "paragraph_or_subsection"), ("####", "letter"), ("#####", "roman"),
]
# Yerel LangChain ortamınızda:
# splitter = MarkdownHeaderTextSplitter(headers_to_split_on=headers_to_split_on,
#                                     strip_headers=False)
```

Bu betik LangChain bağımlılığı yüklemez. Küçük Qwen modelleri için başlıkları chunk içeriğinde tutun. Metadata'ya kaynak hash'i, sayfa, bbox, madde yolu ve blok kimliğini ekleyin; bunlar evidence dosyasından alınabilir. Tabloyu karakter sınırından ortadan bölmeyin: satır gruplarıyla parçalayın ve sütun başlıklarını her parçada taşıyın. Tam tablo ilişkileri için Markdown yanında hücre koordinatlarını kullanın.

## Bilinen sınırlar

- OCR uygulanmaz. Taranmış veya metin katmanı bozuk PDF yerel OCR ve ayrıca görsel doğrulama gerektirir; otomatik tahmin yapılmaz. Kaynak PDF ve sayfa görselleri saklanır.
- Diyagram okları, şekiller ve görüntü içi metin düz Markdown'a kayıpsız aktarılamaz. Bunlar kaynakta korunur, anlamsal Markdown dönüşümü tamamlanmış sayılmaz.
- RTL, karmaşık sütunlar, üst simgeler, dipnot bağlantıları, `(i)` gibi Roma rakamı/harf belirsizlikleri ve belgeye özel başlıklar elle kontrol gerektirir. Hukuki başlık tanıma tüm belge biçimleri için evrensel değildir.
- Satır sonunda `sector-` gibi tireler korunur; kelime birleştirme tahmini yapılmaz. Belirsiz sayfa devamları otomatik birleştirilmez.
- Marj filtresi geometrik bir sezgiseldir. Dışlanan kelimeler kaybolmaz; sidecar'da incelenebilir. Marjdaki gerçek içerik yanlış dışlanmışsa yeni ayarla yeniden çalıştırın.
- Sayfa numarası filtresi son fiziksel satırdaki ortalanmış sayılara uygulanır; kenara hizalı sayılar emin olunamadığı için tutulabilir. Dipnot numaraları ve metni korunur.
- Sayfa metni, görseller ve evidence sayfa sayfa işlenir; tüm metin RAM'de toplanmaz. pdfplumber/pypdf belge nesneleri ve sayfa raporu listesi nedeniyle toplam RAM kesin O(1) değildir. Güvenilmeyen/dev PDF'ler için işletim sistemi düzeyinde RAM, CPU ve süre limitleri bulunan ayrı işlem/container kullanın.
- `verify`, üretilen blok içindeki kelime sırasını kontrol eder; PDF'nin gerçek görsel okuma sırasını matematiksel olarak doğrulamaz. İki kontrol birbirine karıştırılmamalıdır.

## Testler

`test_pipeline.py` kelime kaybı/tekrarı, sıralama değişikliği, metin mutasyonu, Markdown kaçışları, çok satırlı/birleşik tablo hücreleri, geometrik fallback, marj filtresi, hukuki hiyerarşi, dosya değişikliği ve kabul kapısını sentetik verilerle test eder. Testler için harici PDF gerekmez.

## Depo içeriği

Depo yalnızca uygulama kodu, sentetik birim testleri, bağımlılık listesi ve kullanım belgesini içerir. Kaynak PDF'ler, sayfa görselleri, dönüşüm çıktıları ve yerel denetim kayıtları `.gitignore` ile hariç tutulur.
