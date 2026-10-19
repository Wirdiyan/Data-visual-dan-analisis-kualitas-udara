# Air Quality Tree — Final

Versi ini mengembangkan project menjadi aplikasi 3D data storytelling dengan empat lapisan analisis:

1. Current air condition — nilai parameter pada Date + Time yang dipilih.
2. Descriptive / distribution analysis — mean, median, standard deviation, quartile, IQR, histogram, box plot, dan outlier.
3. Correlation analysis — Pearson correlation + scatter plot.
4. Temporal analysis — time-series berdasarkan tanggal aktif.

## Cara menjalankan

1. Ekstrak ZIP.
2. Buka `index.html` melalui browser modern. Jika browser memblokir resource lokal, jalankan folder dengan local server sederhana (misalnya Live Server di VS Code).
3. Klik **Choose Excel File**.
4. Pilih dataset Excel yang memiliki kolom `Date` dan `Time`.
5. Pilih tanggal dan geser slider waktu.
6. Gunakan dropdown pada bagian Statistical Analysis untuk mengganti variabel.

## Catatan metodologi

- Nilai `-200` diperlakukan sebagai missing value dan tidak dimasukkan ke perhitungan statistik.
- Correlation menggunakan Pearson r pada pasangan observasi yang sama-sama valid.
- Outlier menggunakan aturan IQR: di bawah Q1 − 1.5×IQR atau di atas Q3 + 1.5×IQR.
- Pollution Index adalah indeks relatif untuk kebutuhan visual storytelling, bukan AQI resmi.
- Pohon 3D mengubah nilai relatif CO, NOx, NO2, temperature, dan humidity menjadi karakteristik visual.
