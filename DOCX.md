# DOCX Markdown dönüştürücüsü

Python 3.10+ ile tamamen yerel çalışır. Ek paket, Word otomasyonu, yapay zekâ modeli veya ağ bağlantısı gerektirmez. `.docx` içindeki OOXML yapısını doğrudan okur; belge içeriğini talimat olarak yürütmez.

## Çalıştırma

```powershell
python docx_to_markdown.py convert "belge.docx" "output\belge-docx-001"
python docx_to_markdown.py verify "output\belge-docx-001"
```

Çıktı klasörü yeni olmalıdır. Sanal ortam kullanılıyorsa `python` yerine `.\.venv\Scripts\python.exe` yazılabilir. PDF hattına ait `run.ps1` yerine DOCX betiğini doğrudan çalıştırın.

Çıkış kodları: `0` teknik doğrulama veya kabul başarılı; `1` hata; `3` dönüşüm tamamlandı, anlamsal inceleme bekleniyor. Dönüşümün `3` dönmesi beklenen davranıştır. Hata halinde `RUNNING.json` kalır; tamamlanmamış çalışmayı kullanmayın.

## Çıktılar

| Dosya | İçerik |
|---|---|
| `source.docx` | Orijinal belgenin değiştirilmemiş kopyası |
| `candidate.md` | Gövde metni, başlıklar, numaralı listeler, tablolar, dipnotlar ve sonnotlar |
| `supplementary.md` | Üstbilgi/altbilgi, yorumlar, silinmiş/gizli metin, alan talimatları, metin kutusu/denklem metni ve diğer ayrılmış içerik |
| `evidence.jsonl.gz` | Blok, kaynak XML metin kimlikleri, tam metin, hedef, başlık/liste yolu, tablo birleşme bilgisi, Markdown bayt aralıkları ve uyarılar |
| `assets/` | Paket içindeki görsellerin yerel kopyaları; desteklenmeyen medya `.bin` olarak korunur |
| `report.json` | Kaynak/çıktı hash'leri, kaynak metin parçaları ve karakter sayıları, uyarı sayıları |
| `review.html` | Kaynak XML metni ile Markdown'ı blok bazında karşılaştıran çevrimdışı ekran |
| `review-template.json` | Hash'lere bağlı insan inceleme şablonu |

`review.html` dosyasını tarayıcıda açın. Ekran Word sayfa düzeninin görüntüsü değildir. Gerçek görsel karşılaştırma için `source.docx` dosyasını yerel belge okuyucunuzda ayrıca açın. DOCX sayfa sayısı fontlara, yazıcıya ve renderer'a bağlıdır; betik sayfa numarası uydurmaz.

## Doğrulama yaklaşımı

1. ZIP parçaları boyut, toplam açılım, yinelenen yol, dizin aşımı ve sıkıştırma oranı açısından kontrol edilir. XML DTD/entity bildirimleri reddedilir. Harici ilişkiler indirilmez; alanlar, makrolar ve gömülü nesneler çalıştırılmaz.
2. Ayrı bir XML geçişi `w:t`, `w:delText`, `w:instrText`, `w:delInstrText`, `a:t`, `m:t` metinlerini kaynak parça ve sıra kimliğiyle SQLite envanterine yazar.
3. Paragraf ve tablolar özgün XML sırasıyla işlenir. Word'ün bir kelimeyi farklı run'lara bölmesi kelimeye fazladan boşluk eklenmesine yol açmaz. Metne yazım düzeltmesi, Unicode normalizasyonu veya tahmini tire birleştirmesi uygulanmaz.
4. Her tanınan metin parçasının tam içeriği bir kez `candidate.md` veya `supplementary.md` hedefine atanmalıdır. Eksik, değiştirilmiş, yinelenmiş veya sırası değişmiş kaynak kimliği hata verir. Otomatik liste numaraları ve Markdown işaretleri üretilmiş içerik olarak ayrı tutulur.
5. `verify` kaynağı yeniden açar, envanteri yeniden kurar, fragment içeriğini ve sırasını kontrol eder; üretilen Markdown baytlarını ve tüm kayıtlı hash'leri doğrular.

Bu kontroller **tanınan XML metninin muhasebesini** doğrular; Word'ün görsel anlamının kayıpsız Markdown'a aktarıldığını kanıtlamaz. Tanınmayan XML/ikili içerik kaynak DOCX'te korunur ve ek parça uyarısıyla incelemeye bırakılır. Rapor ve hash'ler kriptografik imzalı bir güven sistemi değildir.

## Korunan yapılar ve sınırlar

- Başlıklar: `outlineLvl`, `Heading1…9`, `Title` ve `basedOn` zinciri; ayrıca bağımsız `CHAPTER`, `ANNEX`, `PART`, `TITLE` ve `Article` satırları. Markdown'ın altı başlık seviyesi sınırı geçerlidir.
- Listeler: `numId`, `ilvl`, `abstractNum`, `lvlText`, yaygın sayı/harf/Roma biçimleri, başlangıç değerleri ve `startOverride`; iç içe seviye ve yeniden başlama kuralları. Word numaralandırmasının bütün varyantları desteklenmez; numaralandırılmış bloklar incelemeye işaretlenir. Liste yolu sidecar metadata'da ayrıca korunur.
- Tablolar: satır/hücre sırası, çok paragraflı hücrelerde `<br>`, kaynak `|` karakterlerinin kaçışı. İlk veri satırını yanlış başlık saymamak için `Column 1…` başlıkları üretilir. `gridSpan`/`vMerge` geometrisi kayıtta tutulur; birleşik hücre metni kopyalanmaz. Markdown row/col span desteklemediğinden birleşmeler görsel inceleme gerektirir. İç içe ve sarılmış tablolar metin fallback'iyle korunabilir; ilişkisel yapının tam korunduğu ileri sürülmez.
- Dipnotlar/sonnotlar: referanslar ile kaynak kimlikli bölümler; yorumlar ayrı dosyaya bağlanır. Özel ayraç notları ana metne karıştırılmaz. Özel numara biçimleri, yerleşim ve kırık kaynak referansları ayrıca kontrol edilmelidir.
- Bağlantılar: `http`, `https`, `mailto` hedefleri korunur; dışarıya istek yapılmaz. İç bağlantılar ve desteklenmeyen URL biçimleri metadata'da tutulup işaretlenir.
- Kalın/italik: doğrudan run özellikleri yansıtılır. Tüm Word karakter stilleri, renk, üst/alt simge, şekil konumu ve görsel özellikler Markdown'da temsil edilmez. Üzeri çizili ve gizli metin inceleme gerektirir.
- Değişiklik izleme: eklenen/güncel metin adayda, silinmiş ve taşınmış-eski metin ayrı dosyada tutulur. Word'ün değişiklik kabul işlemi uygulanmaz; revizyonlu belgeler inceleme gerektirir.
- Alanlar: kaydedilmiş görünen sonuç korunur; `PAGE`, `REF`, `TOC` vb. talimatlar ayrı tutulur. Sonuç yeniden hesaplanmaz ve güncelliği garanti edilmez.
- Denklemler, grafikler, metin kutuları, `altChunk`, gömülü nesneler ve özel semboller tam anlamsal dönüşüm kapsamında değildir. Kaynak dosya esas alınır; tanınan metin ayrıca korunur ve uyarı üretilir.
- Eski `.doc`, şifreli Office paketleri ve Strict OOXML namespace desteklenmez. Uygun `.docx` olarak kaydedilmelidir. OCR yapılmaz.

Bellek kullanımı belge metninin tamamını RAM'e toplamadan mantıksal bloklarla sınırlanır. Tek bir büyük tablo ile stiller/numaralandırma tanımları bellekte tutulabilir. Varsayılan limitler XML/paket parçası başına 64 MiB, toplam ZIP açılımı 512 MiB ve 10.000 parçadır.

## İnceleme ve kabul

`review-template.json` dosyasını kopyalayın; `reviewer` alanını ve kontrol edilen her blok için `approved: true`, açıklayıcı `note` değerlerini doldurun. Supplementary içerik ile görseller de inceleme kapsamındadır. Kaynak dosya ile görsel ve anlamsal kontrol yapmadan toplu onay vermeyin.

```powershell
python docx_to_markdown.py release "output\belge-docx-001" --review "output\belge-docx-001\my-review.json"
```

Eksik onay veya dosya değişikliği kabulü engeller. Başarılı kabul `ready.md` üretir. `ready.md`, `candidate.md` içeriğidir; ayrı korunan içerik otomatik olarak ana gövdeye eklenmez. Uygulamanız için dipnot dışındaki yardımcı içerik de aranabilir olmalıysa `supplementary.md` dosyasını **belirgin kaynak türü metadata'sıyla ayrı** indeksleyin; silinmiş metni yürürlükteki hüküm gibi sunmayın. Görseller için `assets/` klasörünü Markdown yanında tutun.

## Testler

```powershell
python -m unittest -v test_docx_pipeline
# PDF bağımlılıkları da kuruluysa her iki hattı birlikte test etmek için:
python -m unittest -v test_docx_pipeline test_pipeline
```

Testler geçici, sentetik OOXML paketleri üretir; kullanıcı belgesi gerektirmez. Başlık mirası, liste yeniden başlatma, tablo hücreleri, revizyonlar, notlar, bağlantılar, görseller, metin kaybı/sırası, çıktı değişikliği, kaynak envanteri ve güvenli paket okuma denetlenir. Bu testler gerçek bir belgenin Word görüntüsüyle yapılmış görsel kabulün yerine geçmez.
