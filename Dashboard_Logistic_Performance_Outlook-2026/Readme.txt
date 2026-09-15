canva:
-Performance against budget
value (budget and forecast)
-persentase donut chart (teks persentase)
1. Kartu Metrik Utama (Top KPI Cards)
Total Shipment (21.40 Mt)

Logika Measure: SUM(Tonnage) dibagi 1 Juta.

Sumber Data: Sheet graph 3, Baris 6, Kolom D - M (Jan-Oct).

Demurrage (0.36 $/Ton)

Logika Measure: [Total Demurrage $] / [Total Shipment].

Sumber Data: Sheet graph 3, Baris 5, Kolom D - M (Rata-rata).

Penalty (+0.32 $/Ton)

Logika Measure: [Total Penalty $] / [Total Shipment].

Sumber Data: Sheet graph 3, Baris 13, Kolom D - M (Rata-rata).

Net Logistics Impact (-0.89 MUSD)

Logika Measure: [Total Penalty $] - [Total Demurrage $].

Sumber Data: Kalkulasi Measure (Gabungan nilai total Penalty dikurangi total Demurrage).

2. Grafik Keuangan (Financial Charts)
Financial Impact (Total Demurrage 7.81 M)

Logika Measure: SUM(BoCT $) + SUM(Non BoCT $) kemudian di-absolut-kan.

Sumber Data: Sheet present dem, Baris 10 - 19, Kolom G (BoCT $) & Kolom J (Non BoCT $).

Financial Impact (Penalty Benefit 6.92 M)

Logika Measure: SUM(Tonnage * Penalty Rate) per bulan.

Sumber Data: Sheet graph 3, Perkalian Baris 6 & Baris 13, Kolom D - M.

Bonus Penalty (Grafik Batang Hijau)

Logika Measure: Tonnage * Penalty Rate (Difilter per bulan).

Sumber Data: Sheet graph 3, Perkalian Baris 6 & Baris 13, Kolom D - M.

3. Grafik Performa & Kontribusi (Performance Charts)
Monthly Average Demurrage (Garis Biru/Hijau)

Logika Measure: Nilai dari baris Actual/Outlook Demurrage per bulan.

Sumber Data: Sheet graph 3, Baris 5, Kolom D - M.

Performance Against Budget (Demurrage Budget 0.50)

Logika Measure: Diambil dari nilai Plan/Budget Demurrage (dibulatkan).

Sumber Data: Sheet graph 3, Baris 4, Kolom D - M.

Performance Against Budget (Penalty Budget -0.10)

Logika Measure: Diambil dari nilai Plan Penalty.

Sumber Data: Sheet graph 3, Baris 12, Kolom D - M.

Demurrage Cost Contribution (Donut BoCT - 4.38M)

Logika Measure: SUM(BoCT Demurrage $).

Sumber Data: Sheet present dem, Baris 10 - 19, Kolom G.

Demurrage Cost Contribution (Donut Non-BoCT - 3.43M)

Logika Measure: SUM(Non BoCT Demurrage $).

Sumber Data: Sheet present dem, Baris 10 - 19, Kolom J.

Throughput Realization (Actual Jan-Jun)

Logika Measure: SUM(Tonnage) difilter khusus bulan Januari hingga Juni.

Sumber Data: Sheet graph 3, Baris 6, Kolom D - I.

Throughput Realization (Forecast Jul-Oct)

Logika Measure: SUM(Tonnage) difilter khusus bulan Juli hingga Oktober.

Sumber Data: Sheet graph 3, Baris 6, Kolom J - M.