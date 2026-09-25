# 📅 Sınav Takvimi Optimizasyon Sistemi

Eğitim kurumlarındaki karmaşık sınav planlama süreçlerini otomatize etmek, insan kaynaklı hataları önlemek ve çakışmasız bir takvim oluşturmak amacıyla geliştirdiğim Python tabanlı bir çözümdür. 

## 🚀 Projenin Amacı ve Çözdüğü Sorunlar
Sınav takvimlerini manuel olarak hazırlamak; öğrenci çakışmaları, gözetmen müsaitliği ve derslik kapasitelerinin ayarlanması gibi operasyonel zorluklar içerir. Bu proje, 3 temel Excel dosyasını (Sınav Listesi, Öğrenci Kayıtları, Oda/Zaman Tablosu) girdi olarak alır ve "sıfır çakışma" garantisiyle otomatik bir program üretir.

## ✨ Temel Özellikler
*   **Üçlü Çakışma Kontrolü:** Öğrenci, gözetmen ve derslik bazlı zaman çakışmalarını aynı anda engeller.
*   **Kapasite Optimizasyonu:** Derslikleri, gerçek kapasitelerinin 1/3'ü oranında (sınav düzeni) kullanarak atama yapar.
*   **Önceliklendirme Algoritması:** "Ortak" havuz derslerini tespit edip sıralamada öne alarak yerleştirme zorluğunu en aza indirir.
*   **Web Arayüzü:** Herhangi bir kurulum gerektirmeyen, yerel sunucu (localhost) üzerinden çalışan HTML tabanlı kullanıcı dostu bir arayüz sunar.
*   **Gelişmiş Raporlama:** `openpyxl` kütüphanesi kullanılarak; ortak derslerin sarı renkle vurgulandığı, sütun genişliklerinin otomatik ayarlandığı, 3 farklı sekmeden (Tüm Program, Ortaklar, Bölüm Dersleri) oluşan profesyonel bir Excel çıktısı üretir.

## 🛠️ Kullanılan Teknolojiler
*   **Python:** Temel algoritma ve yerel sunucu (HTTP Server) mimarisi.
*   **Pandas:** Excel veri manipülasyonu, sıralama ve filtreleme işlemleri.
*   **Openpyxl:** Çıktı dosyasının dinamik hücre biçimlendirmesi ve renklendirilmesi.

## 📌 Nasıl Çalıştırılır?
1. Repoyu klonlayın ve proje dizinine gidin.
2. Terminal üzerinden `python program.py` komutunu çalıştırın.
3. Tarayıcınızda `http://localhost:1903` adresine giderek arayüze ulaşın.
