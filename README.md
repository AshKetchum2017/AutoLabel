# AutoLabel

AutoLabel adalah macro VBA untuk CorelDRAW yang membuat label teks pada dokumen aktif. Label ditempatkan pada `Layer 1` di page aktif atau Master Page, dengan posisi yang menyesuaikan mode dan objek penanda yang tersedia.

Repository ini memuat dua varian UserForm: **Auto Label** (`AutoLabelWizard`) dan **Auto Label BP** (`AutoLabelWizardBP`). Keduanya mendukung format teks, pengambilan nama pelanggan dan tanggal dari folder dokumen CDR, serta integrasi dengan `MacroRunner`. Varian BP menambahkan pilihan operator untuk pewarnaan nama pelanggan.

## Main Features

### Membuat dan memosisikan label

Isi `txbLabel`, pilih salah satu mode, lalu tekan **Submit** (`cmdSubmit`). Macro membuat artistic text pada `Layer 1`; bila layer belum tersedia, macro membuatnya. `chkMasterPage` menentukan apakah label masuk ke Master Page atau page aktif.

| Mode | Kontrol | Posisi dan perilaku |
| --- | --- | --- |
| Uni Label | `optUniLabel` | Mengikuti objek bernama `filename` pada layer `scpro2_printonly` jika ada; teks rata kanan. |
| Kiss Label | `optKissLabel` | Mengikuti penanda `filename`; teks rata kiri dan ukuran font 4 pt. Jika `/*~` membaca folder PART bernomor, tambahkan baris `PT.n`. |
| Hi Die | `optHiDie` | Mengikuti penanda persegi panjang yang sesuai pada layer `Regmark`; teks rata kiri. Nama pelanggan diberi kapital awal tiap kata, kecuali jika `chkUppercase` aktif. |
| Hi Kiss | `optHiKiss` | Mengikuti penanda persegi 3 × 3 mm pada `Regmark`; teks rata kiri dan diputar. Arah putaran Auto Label dan Auto Label BP berbeda. |

Jika penanda untuk mode terpilih tidak ditemukan, kode memakai koordinat bawaan mode tersebut. Periksa hasilnya pada layout dokumen sebelum produksi.

### Format teks dan konfigurasi

| Kontrol | Fungsi |
| --- | --- |
| `cmbFontSize` | Pilihan ukuran 4–24 pt. `optKissLabel` selalu memakai 4 pt saat label dibuat. |
| `chkBold`, `chkItalic` | Mengatur ketebalan dan kemiringan teks. |
| `chkUppercase` | Mengubah seluruh teks utama menjadi huruf kapital; `optKissLabel` juga menerapkan kapital secara bawaan. |
| `chkSmallCaps` | Menerapkan small caps pada bagian teks yang dikenali. |
| `chkColorize` | Mewarnai bagian teks menurut pola label. Pada varian BP, warna pelanggan mengikuti operator bila pola tanggalnya dikenali. |
| `chkFraction` | Menerapkan fitur OpenType fraction pada pecahan yang dikenali dan memakai `/` untuk token tanggal gabungan. |
| `chkEmDash` | Mengubah kemunculan pertama `" - "` menjadi em dash dan memakai `-` untuk token tanggal gabungan. |
| `chkMasterPage` | Menaruh label pada `Layer 1` milik Master Page; jika tidak dipilih, gunakan page aktif. |

Tekan **Save Config** (`cmdSaveConfig`) untuk menyimpan pilihan kontrol dan ukuran font ke pengaturan VBA. Pilihan Auto Label dan Auto Label BP disimpan pada section terpisah, masing-masing `AutoLabel` dan `AutoLabelBP` di bawah `RinCorelMacros`. Teks `txbLabel` tidak ikut disimpan sebagai konfigurasi.

### Operator pada Auto Label BP

`AutoLabelWizardBP` menyediakan `cmbOperator` dengan pilihan `Saiful`, `Hafizh`, `Dori`, dan `Reni`. Saat `chkColorize` aktif dan bagian akhir label cocok dengan pola tanggal yang dikenali, nama pelanggan diberi warna CMYK berikut:

| Operator | Warna nama pelanggan (C/M/Y/K) |
| --- | --- |
| `Saiful` | `0/100/100/0` |
| `Hafizh` | `0/0/0/100` |
| `Dori` | `0/100/0/0` |
| `Reni` | `100/100/0/0` |

Pilihan operator dan opsi BP lainnya dapat disimpan dengan `cmdSaveConfig`. Pemetaan warna tersebut berlaku pada varian BP; format warna pada Auto Label biasa mengikuti aturan teksnya sendiri.

## Token label dari lokasi CDR

Token dapat diletakkan di awal, tengah, atau akhir `txbLabel` dan digabung dengan teks biasa. Sumbernya adalah **folder dokumen CDR aktif**, bukan tanggal komputer.

| Token | Hasil |
| --- | --- |
| `/*~` | Nama folder pelanggan dalam huruf kecil. Jika CDR berada di folder `PART n`, ambil nama folder induknya. |
| `/*dd` | Tanggal sumber dua digit. |
| `/*mm` | Bulan sumber dua digit. |
| `/*yy` | Dua digit terakhir tahun sumber. |
| `/*ddmm` | Tanggal dan bulan. |
| `/*ddmmyy` | Tanggal, bulan, dan dua digit tahun. |

Misalnya CDR berada di `D:\Produksi\2026\09. SEPTEMBER\28\Contoh Printing\PART 003\Desain.cdr`. Untuk `txbLabel` berisi `/*~ - /*ddmmyy`, hasil sebelum perubahan huruf dan em dash adalah `contoh printing - 280926`. Bila `chkFraction=True`, bagian tanggal menjadi `28/09/26`; bila `chkEmDash=True` dan fraction tidak aktif, tanggal menjadi `28-09-26`. Saat keduanya aktif, `chkFraction` didahulukan untuk pemisah tanggal.

Pada contoh tersebut, `optKissLabel` menghasilkan teks utama dalam huruf kapital dan menambahkan baris `PT.3` karena `/*~` membaca folder `PART 003`. Token tanggal saja tidak menambahkan baris PART.

Token tanggal mencari folder hari `1`–`31` di dalam folder bulan dengan format seperti `09. SEPTEMBER`, lalu mencari tahun pada folder leluhurnya. Token `/*~` memerlukan folder pelanggan yang valid. Jika struktur yang diperlukan tidak ditemukan, atau token `/*` tidak dikenal, proses menampilkan error. Simpan CDR di struktur folder yang sesuai sebelum memakai token.

## Alur penggunaan

1. Buka dokumen di CorelDRAW; simpan sebagai CDR di folder yang sesuai jika akan memakai token.
2. Buka `AutoLabelWizard` atau `AutoLabelWizardBP`.
3. Isi `txbLabel`, pilih mode, atur format teks dan `chkMasterPage`. Pada BP, pilih `cmbOperator`.
4. Tekan `cmdSaveConfig` jika ingin memakai lagi opsi tersebut pada sesi berikutnya.
5. Tekan `cmdSubmit`. Periksa posisi, isi, dan warna label pada dokumen.

**Submit** menutup UserForm setelah label berhasil dibuat. Jika proses gagal, form tetap terbuka agar input dapat diperiksa. Macro mengembalikan unit dokumen ke unit sebelumnya setelah proses atau saat menangani error.

## Integrasi MacroRunner

MacroRunner dapat membuka kedua varian dan menjalankan **MacroBehavior**. Gunakan pasangan MacroID dan nama form yang sesuai:

| MacroID | Nama form dalam Behavior |
| --- | --- |
| `AutoLabel` | `AutoLabelWizard` |
| `AutoLabelBP` | `AutoLabelWizardBP` |

Contoh Auto Label dengan pengaturan form yang telah tersimpan:

```text
{AutoLabel:
    AutoLabelWizard[
        txbLabel="/*~";
        @cmdSubmit
    ]
}
```

Contoh Auto Label BP dengan operator dan token tanggal:

```text
{AutoLabelBP:
    AutoLabelWizardBP[
        txbLabel="/*~ - /*ddmmyy";
        cmbOperator="Saiful";
        optKissLabel=True;
        @cmdSubmit
    ]
}
```

`txbLabel` wajib diisi dengan string nonkosong dalam Behavior, dan `@cmdSubmit` harus menjadi instruksi terakhir. Satu blok hanya boleh mempunyai satu tahap form. Opsi yang tidak ditulis tetap memakai nilai awal yang dimuat UserForm, termasuk konfigurasi tersimpan. `Default` mengembalikan nilai awal kontrol yang didukung; `Nothing`, `Empty`, dan `Null` melewati assignment. `cmbOperator` hanya tersedia untuk `AutoLabelWizardBP`, dengan empat nama operator di atas; `cmbFontSize` menerima angka 4–24.

`MRTargetBridge.ValidateBehavior` memeriksa grammar dan kontrak target sebelum antrean dijalankan. Ketersediaan dokumen, struktur folder untuk token, dan hasil pembuatan label baru diketahui saat `@cmdSubmit` berjalan. Jika Submit gagal, kesalahan diteruskan ke MacroRunner dan form dibiarkan terlihat untuk pemeriksaan manual. Saat form ditutup setelah Submit berhasil, callback memberi tahu MacroRunner bahwa langkah selesai.

## Source dan pemasangan

Repository menyediakan source VBA dan code-behind form, bukan file desain UserForm lengkap. Pada project GMS masing-masing, siapkan UserForm dengan `(Name)` yang tepat dan kontrol bernama seperti di source.

| Lokasi | Peran |
| --- | --- |
| `src/forms/AutoLabel.vba` | Code-behind `AutoLabelWizard`. |
| `src/forms/AutoLabelBP.vba` | Code-behind `AutoLabelWizardBP`. |
| `src/classes/ALPresenter.cls`, `ALInputModel.cls`, `ALLabelOptions.cls` | Pembuatan dan posisi label, parsing token, serta snapshot opsi. |
| `src/classes/ALBehaviorContract.cls`, `ALBehaviorSession.cls`, `MRBehavior*.cls` | Validasi dan eksekusi Behavior. |
| `src/modules/MRTargetBridge.bas` | Bridge untuk membuka form serta memvalidasi dan menjalankan Behavior dari MacroRunner. |

Varian biasa dan BP memakai class bersama; sertakan class yang diperlukan dalam masing-masing project GMS yang digunakan. Nama kontrol, layer, dan penanda perlu cocok dengan code-behind. Integrasi MacroRunner tidak memerlukan reference langsung ke project VBA MacroRunner.

## Lisensi dan masukan

Project ini menggunakan lisensi MIT; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
