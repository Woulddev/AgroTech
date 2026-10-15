# Priva VP9508 Greenhouse Sensor Dashboard

## Haqqında
Bu vebsayt 6+ il (2020-2026) ərzindəki Priva VP9508 istixana iqlim nəzarət sisteminin sensor datalarını analiz edir.

## Texniki Detallar
- **Sistem**: Priva Integro / Connext VP9508
- **Fayl formatı**: Binary .dlt (indeks) + .dof (verilənlər) - zlib sıxılmış
- **Rekord strukturu**: int32(unix_timestamp) + double(value) = 12 bayt
- **Kanallar**: 371 unikal sensor kanali
- **Tarix aralığı**: 25 Yanvar 2020 → 25 Aprel 2026 (588 gün)
- **Ümumi ölçümlər**: ~6.5 milyon rekord
- **Interval**: ~5 dəqiqəlik ölçmə intervalı

## Sensor Növləri
- Xarici hava: temperatur, külək, radiasiya, yağış, don
- İstixana: temperatur (4 bölmə), rütubət (4 bölmə)
- Isıtma: qazan temperaturu, boru temperaturu, bufer dolumu
- Ventilyasiya: açılma faizi, küləyin istiqaməti
- CO₂: konsentrasiya, dozaj nəzarəti
- Su: su temperaturu, EC, pH, axın sürəti
- Enerji: qaz sərfi, ümumi qaz

## Necə işlətmək
```bash
pip install flask
python3 run_dashboard.py
# Brauzer: http://localhost:5000
```

## Yeni data axını
- Frontend ümumi kartlar və mini qrafiklər üçün `static/data.json` faylındakı mövcud `daily/monthly` datasını istifadə edir.
- Detallı qrafikdə 1 günlük aralıq seçiləndə backend `greenhouse_data_clean.pkl` faylındakı xam timestamp ölçmələrini saatlıq ortalamaya çevirir və `00:00`-`23:00` aralığını qaytarır.
- Manual/planlanan overlay datası real sensor datasından ayrı `manual_data.json` faylında saxlanılır.

## Admin panel
- Sağ yuxarıdakı `Admin Panel` düyməsi backend session login istifadə edir.
- Şifrə `ADMIN_PASSWORD` environment dəyişəni ilə təyin oluna bilər.
- Lokal default şifrə: `admin123`.
- Admin panel manual overlay datası ilə yanaşı threshold/alert qaydalarını da idarə edir.

## İstifadəçi interfeysi
- Light/Dark mode sağ yuxarıdakı tema düyməsi ilə dəyişir və seçim brauzerdə saxlanılır.
- Global filter paneli tarix aralığı, sensor qrupu, sensor axtarışı və manual overlay aktiv/deaktiv idarəsini təmin edir.
- Əsas qrafik və detal modalı CSV/PNG export dəstəkləyir.
- Qrafiklərdə mouse wheel ilə horizontal zoom, `Zoom reset` ilə ilkin görünüşə qayıtmaq mümkündür.
- Data keyfiyyəti paneli seçilmiş dövr üçün boş data, real data coverage və manual overlay sayını göstərir.
- Alert paneli admin tərəfindən təyin edilən threshold-lara əsasən seçilmiş dövrdə limitdən kənara çıxan dəyərləri göstərir.

## Manual import
Admin panel CSV, JSON və Excel (`.xlsx`, `.xlsm`) import qəbul edir. Importdan əvvəl `Preview` ilə səhv sətrləri yoxlamaq, `Template CSV` ilə nümunə fayl endirmək olar. Gözlənilən sütunlar:

```text
channel_id,date,time,value,enabled
0,2026-04-24,01:00,25,true
```
