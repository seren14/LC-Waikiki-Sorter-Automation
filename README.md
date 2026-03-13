# LC Waikiki Depo Sorter Otomasyon Sistemi

LC Waikiki Yalova deposu sorter departmanı için geliştirilen kapsamlı Excel VBA macro otomasyon sistemi.

## Proje Hakkında

Bu sistem, LC Waikiki Yalova deposundaki sorter operasyonlarının raporlanmasını ve takibini otomatikleştirmek amacıyla geliştirilmiştir. Sistem aktif olarak operasyonel kullanımdadır.

## Özellikler

- **Performans Dashboard'ları:** TTY ve TTS sorter lokasyonlarından otomatik performans raporları
- **Hata Analizi (Hata Analizi):** GL bazlı hata toplamları ve detaylı analiz raporları
- **Koli Kapama Raporları:** Koli kapama süreçlerinin takibi ve raporlanması
- **Sorter No-Info Smart Match (v2):** NI (No-Info) sorter item'larını GL görev listeleriyle çapraz referanslayarak eşleştiren akıllı modül
- **Vardiya Bazlı Filtreleme:** Vardiya bilgisine göre veri ayrıştırma
- **Toplu/Bireysel Drop Sınıflandırması:** Packet seviyesinde GL intersection mantığı

## Teknik Detaylar

- **Dil:** Excel VBA
- **Veri Yapıları:** Nested Dictionary yapıları (cross-GL veri kontaminasyonunu önlemek için)
- **Raporlama:** Pivot tablo bazlı otomatik raporlar
- **Kapsam:** 2 sorter lokasyonu (TTY, TTS)

## Kullanım

1. Excel dosyasını açın
2. Sorter verilerini ilgili sayfaya yapıştırın
3. Macro'yu çalıştırın
4. Otomatik oluşturulan raporları inceleyin

## Geliştirici

**Semih Eren Halgalı**
- LC Waikiki Yalova Deposu - Sorter Departmanı
