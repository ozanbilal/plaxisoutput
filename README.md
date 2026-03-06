# PLAXIS Output Export Tools

Bu repo, PLAXIS Output uzerinden veri cekmek ve toplu rapor uretmek icin kullanilan araclari icerir.

## Icerik

- `plaxis_export_gui.py`
  Tkinter tabanli arayuz. Faz yukleme, structural moment export, node spectrum export ve secili node time history export seceneklerini icerir.

- `export_plaxis_data.py`
  Ana is mantigi. PLAXIS Output API baglantisi, veri cekme, Excel yazma, PNG grafik uretme ve chart olusturma burada bulunur.

- `run_plaxis_multiphase_cli.py`
  Terminalden calistirmak icin CLI wrapper.

## Desteklenen akislar

### 1. Structural Moment Analysis

- Fazlari API uzerinden yukler
- X ve Y yonleri icin ayri secim yapar
- EmbeddedBeam ve Plate gruplarini okur
- `MEnvelopeMax2D` ve `MEnvelopeMin2D` degerlerini alir
- Excel wide sheet ve Excel icinde chart uretir
- PNG grafikler de yazar

### 2. Node Spectrum Analysis

- Secili CurvePoint ler icin faz bazli ivme-zaman serisi okur
- Spektrum hesaplar
- X ve Y yonleri icin ayri overlay ve mean grafikler uretir
- Excel wide sheet + Excel chart + PNG ciktilari olusturur
- Istenirse her faz icin node time history CSV dosyalarini alt klasore yazar

## Gereksinimler

- PLAXIS Output remote scripting acik olmali
- Python 3
- Gerekli paketler:
  - `numpy`
  - `pandas`
  - `openpyxl`
  - `matplotlib`
  - `plxscripting`
  - `pywinauto` (sadece GUI points export akislarinda gerekir)

## GUI kullanimi

```powershell
python plaxis_export_gui.py
```

Arayuzde:

- `Load Phases` ile fazlari cek
- X ve Y listelerinden faz sec
- Gerekirse `Load Structural Objects` ve `Load CurvePoints` calistir
- `PNG DPI` ile grafik cozunurlugunu ayarla
- `Save node time histories into phase subfolders` secenegi ile faz bazli CSV export ac

Node time history alt klasor secenegi aciksa, workbook yanina su yapida dosyalar yazilir:

```text
<output>_time_history/
  DD2_X_20030501002708_1201_H2/
    Node_24388_22_90_20_95.csv
```

Klasor adi faz gorunen isminden uretilir; `[Phase_6]` gibi kisim kullanilmaz.

## CLI kullanimi

### Node export

```powershell
python run_plaxis_multiphase_cli.py node ^
  --host localhost ^
  --port 10000 ^
  --password "YOUR_PASSWORD" ^
  --out "C:\\temp\\node_results.xlsx" ^
  --plot-dpi 200 ^
  --save-node-timehistory-subfolders
```

### Structural export

```powershell
python run_plaxis_multiphase_cli.py structural ^
  --host localhost ^
  --port 10000 ^
  --password "YOUR_PASSWORD" ^
  --out "C:\\temp\\structural_results.xlsx" ^
  --plot-dpi 200
```

## Output yapisi

### Structural workbook

- `Phases`
- `Selections`
- `MomentRawLong`
- `MomentAvgByDir`
- `MomentWide_*`
- `_Status`

### Node workbook

- `Phases`
- `Selections`
- `NodeTimeHistoryLong`
- `NodeSpectrumLong`
- `NodeSpectrumMean`
- `Spec_*`
- `SpecPhase_*`
- `SpecMean_*`
- `_Status`

## Notlar

- Excel icindeki chart lar numeric axis kullandigi icin eksen araliklari otomatik hesaplanir.
- Structural wide tablolar obje bazli blok kolonlar halinde yazilir; bu sayede veriler daginik gorunmez.
- Grafik legend lari daha kompakt yerlesimle uretilir.
