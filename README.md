# 🏥 EÜTF Sınav Analiz ve Rapor Sistemi

Ege Üniversitesi Tıp Fakültesi — Tıp Eğitimi Anabilim Dalı  
Ölçme ve Değerlendirme Komisyonu

## Özellikler

- **BYS / SBYS uyumlu** — Blok Yazılı Sınav ve Staj Bloğu sınavlarını destekler
- **Gelişmiş psikometri** — KR-20, Ferguson's δ, Guttman Split-Half, SEM, kaliteli alt küme KR-20
- **Karar destek matrisi** — Güçlük × Ayırt edicilik çapraz tablosu ile otomatik aksiyon önerisi
- **Kesme puanı simülasyonu** — Sorunlu maddeler çıkarıldığında başarı oranı değişim analizi
- **AI değerlendirme** — Gemini ile otomatik psikometrik yorum (opsiyonel)
- **DOCX rapor** — Profesyonel Word çıktısı, tüm grafikler ve tablolar dahil
- **100 üzerinden normalizasyon** — Soru sayısından bağımsız çalışır

## Kurulum

### Streamlit Community Cloud (Önerilen)

1. Bu repo'yu fork edin
2. [share.streamlit.io](https://share.streamlit.io) adresine gidin
3. "New app" → repo'nuzu seçin → `app.py` → Deploy

### Lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Gemini API Key

AI değerlendirme özelliği için:

1. [Google AI Studio](https://aistudio.google.com/apikey) adresinden API key alın
2. Streamlit Cloud'da: Settings → Secrets → aşağıdaki formatı ekleyin:

```toml
GEMINI_KEY = "AIza..."
```

3. Lokal kullanımda: `.streamlit/secrets.toml` dosyası oluşturun (git'e eklemeyin)

## Girdi Dosyaları

Sistem iki Excel dosyası bekler:

| Dosya | Açıklama |
|-------|----------|
| **Öğrenci Soru Analizi** | Öğrenci × soru matrisi (0/1). Header: q1, q2, ..., qN |
| **Frekans Analizi** | Madde bazlı istatistikler. Sütunlar: #, Soru Sahibi, Seçenekler, Zorluk, Ayırt Edicilik |

Her iki dosya da `.xls` veya `.xlsx` formatında olabilir.

## Lisans

Bu yazılım Ege Üniversitesi Tıp Fakültesi Tıp Eğitimi Anabilim Dalı tarafından geliştirilmiştir.
