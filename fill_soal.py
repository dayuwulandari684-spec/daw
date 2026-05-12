from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
import lxml.etree as etree

# ============================================================
# DATA SOAL
# ============================================================

NAMA_PENULIS = "Roseno Afandi"
MAPEL = "Komputer Akuntansi"
KELAS = "XI"
REF = "Modul Komputer Akuntansi Accurate Online Kelas XI, Direktorat PSMK Kemdikbud 2023"

# 20 Soal Pilihan Ganda
pg_data = [
    {
        "no": 1,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pengenalan Accurate Online",
        "cp": "Peserta didik mampu menjelaskan karakteristik dan keunggulan aplikasi Accurate Online sebagai software akuntansi berbasis cloud",
        "ipk": "Mengidentifikasi keunggulan Accurate Online berbasis cloud dibandingkan software akuntansi konvensional",
        "indikator": "Disajikan informasi tentang karakteristik Accurate Online, peserta didik dapat menentukan keunggulan utamanya",
        "tingkat": "Rendah",
        "nilai": "5",
        "soal": "Accurate Online adalah aplikasi akuntansi berbasis cloud yang dikembangkan oleh PT Cipta Piranti Sejahtera. Accurate Online dapat diakses melalui berbagai perangkat tanpa perlu instalasi di setiap komputer pengguna. Berdasarkan karakteristik tersebut, salah satu keunggulan utama Accurate Online dibandingkan software akuntansi konvensional adalah ...",
        "A": "Memerlukan instalasi di setiap komputer yang digunakan",
        "B": "Data tersimpan secara lokal sehingga lebih aman dari serangan internet",
        "C": "Dapat diakses kapan saja dan di mana saja menggunakan koneksi internet",
        "D": "Hanya bisa digunakan oleh satu pengguna dalam satu waktu",
        "E": "Tidak memerlukan koneksi internet sama sekali",
        "jawaban": "C",
    },
    {
        "no": 2,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Setup Data Awal Perusahaan",
        "cp": "Peserta didik mampu melakukan setup data awal perusahaan di Accurate Online secara benar",
        "ipk": "Melakukan prosedur pembuatan database perusahaan baru di Accurate Online",
        "indikator": "Diberikan skenario pembuatan perusahaan baru, peserta didik dapat mengidentifikasi informasi yang diperlukan dalam setup awal",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Saat pertama kali menggunakan Accurate Online untuk membuat database perusahaan baru, pengguna harus mengisi berbagai informasi dasar. Informasi ini akan menjadi landasan dari semua transaksi yang akan dicatat. Berikut ini yang BUKAN merupakan informasi yang diisi pada tahap setup awal perusahaan di Accurate Online adalah ...",
        "A": "Nama perusahaan dan alamat perusahaan",
        "B": "Periode fiskal (tahun buku) perusahaan",
        "C": "Metode pencatatan persediaan (FIFO/Average)",
        "D": "Nama dan biodata lengkap direktur utama perusahaan",
        "E": "Mata uang fungsional yang digunakan",
        "jawaban": "D",
    },
    {
        "no": 3,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Daftar Akun (Chart of Accounts)",
        "cp": "Peserta didik mampu membuat dan mengelola daftar akun di Accurate Online",
        "ipk": "Mengidentifikasi pengelompokan akun dalam Accurate Online berdasarkan kode akun",
        "indikator": "Diberikan kode akun tertentu, peserta didik dapat menentukan kelompok akun yang dimaksud",
        "tingkat": "Rendah",
        "nilai": "5",
        "soal": "Dalam Accurate Online, daftar akun yang digunakan untuk mencatat transaksi keuangan perusahaan dikelompokkan berdasarkan kode angka yang sistematis. Setiap kelompok akun memiliki kode awal yang berbeda. Jika kode akun dimulai dengan angka 1, maka kelompok akun tersebut termasuk dalam kelompok ...",
        "A": "Liabilitas (Kewajiban)",
        "B": "Ekuitas (Modal)",
        "C": "Pendapatan",
        "D": "Beban",
        "E": "Aset (Harta)",
        "jawaban": "E",
    },
    {
        "no": 4,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Saldo Awal Akun",
        "cp": "Peserta didik mampu menginput saldo awal akun di Accurate Online untuk perusahaan yang sudah berjalan",
        "ipk": "Menginput saldo awal seluruh akun neraca di Accurate Online",
        "indikator": "Diberikan kasus perusahaan yang akan mulai menggunakan Accurate Online, peserta didik dapat menentukan menu yang tepat untuk input saldo awal",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Sebuah perusahaan dagang yang telah berjalan selama 2 tahun akan mulai menggunakan Accurate Online. Sebelum mencatat transaksi baru, perusahaan harus memasukkan saldo awal untuk seluruh akun yang dimiliki. Di Accurate Online, menu yang digunakan untuk menginput saldo awal akun-akun neraca adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Buku Besar → Saldo Awal",
        "C": "Menu Pembelian → Faktur Pembelian",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Laporan → Neraca",
        "jawaban": "B",
    },
    {
        "no": 5,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Data Master Pelanggan",
        "cp": "Peserta didik mampu membuat dan mengelola data master pelanggan di Accurate Online",
        "ipk": "Membuat data master pelanggan (customer) di Accurate Online sebelum mencatat transaksi penjualan",
        "indikator": "Diberikan skenario transaksi dengan pelanggan baru, peserta didik dapat mengidentifikasi menu yang tepat untuk input data pelanggan",
        "tingkat": "Rendah",
        "nilai": "5",
        "soal": "Sebelum mencatat transaksi penjualan kepada pelanggan baru, seorang akuntan harus terlebih dahulu membuat data master pelanggan tersebut. Tanpa data master yang lengkap, transaksi tidak dapat dicatat dengan benar. Di Accurate Online, data master pelanggan (customer) disimpan dan dikelola melalui menu ...",
        "A": "Menu Pembelian → Data Pemasok",
        "B": "Menu Persediaan → Data Barang",
        "C": "Menu Penjualan → Data Pelanggan",
        "D": "Menu Buku Besar → Daftar Akun",
        "E": "Menu Laporan → Laporan Penjualan",
        "jawaban": "C",
    },
    {
        "no": 6,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Transaksi Penjualan",
        "cp": "Peserta didik mampu mencatat transaksi penjualan kredit di Accurate Online dengan benar",
        "ipk": "Mencatat faktur penjualan kredit di Accurate Online",
        "indikator": "Diberikan data transaksi penjualan kredit, peserta didik dapat menentukan menu yang tepat untuk mencatatnya",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "PT Maju Jaya melakukan penjualan barang dagangan kepada CV Berkah senilai Rp10.000.000 secara kredit dengan syarat pembayaran Net 30. Akuntan PT Maju Jaya harus mencatat transaksi ini di Accurate Online. Menu yang tepat digunakan untuk mencatat transaksi penjualan kredit tersebut adalah ...",
        "A": "Menu Pembelian → Purchase Order",
        "B": "Menu Penjualan → Faktur Penjualan",
        "C": "Menu Kas & Bank → Penerimaan Kas",
        "D": "Menu Buku Besar → Jurnal Umum",
        "E": "Menu Persediaan → Transfer Stok",
        "jawaban": "B",
    },
    {
        "no": 7,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Data Master Barang / Persediaan",
        "cp": "Peserta didik mampu membuat data master barang dan mengelola persediaan di Accurate Online",
        "ipk": "Membuat data master barang (item list) di Accurate Online",
        "indikator": "Diberikan kondisi sebelum transaksi penjualan/pembelian, peserta didik dapat menentukan menu input data barang",
        "tingkat": "Rendah",
        "nilai": "5",
        "soal": "Sebelum mencatat transaksi penjualan atau pembelian barang, akuntan harus memastikan data master barang tersebut sudah tersedia di sistem. Tanpa data master barang, sistem tidak dapat memproses transaksi dengan baik. Di Accurate Online, untuk membuat data master barang (item list) dilakukan melalui ...",
        "A": "Menu Penjualan → Data Pelanggan",
        "B": "Menu Pembelian → Data Pemasok",
        "C": "Menu Persediaan → Data Barang",
        "D": "Menu Buku Besar → Daftar Akun",
        "E": "Menu Laporan → Laporan Persediaan",
        "jawaban": "C",
    },
    {
        "no": 8,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Transaksi Pembelian",
        "cp": "Peserta didik mampu mencatat transaksi pembelian kredit di Accurate Online dengan benar",
        "ipk": "Mencatat faktur pembelian kredit dari pemasok di Accurate Online",
        "indikator": "Diberikan data transaksi pembelian kredit, peserta didik dapat menentukan menu yang tepat untuk mencatatnya",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "CV Berkah Sejahtera melakukan pembelian bahan baku dari PT Supplier Jaya secara kredit senilai Rp15.000.000 dengan syarat pembayaran 2/10 n/30. Untuk mencatat transaksi pembelian kredit ini dengan benar di Accurate Online, menu yang harus digunakan adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Kas & Bank → Pengeluaran Kas",
        "C": "Menu Pembelian → Faktur Pembelian",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Buku Besar → Jurnal Umum",
        "jawaban": "C",
    },
    {
        "no": 9,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Penjualan",
        "cp": "Peserta didik mampu mencatat transaksi retur penjualan di Accurate Online",
        "ipk": "Mencatat dokumen retur penjualan di Accurate Online sesuai referensi faktur aslinya",
        "indikator": "Diberikan skenario pengembalian barang dari pelanggan, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Pelanggan CV Makmur Abadi mengembalikan 5 unit barang yang telah dibeli dari PT Jaya Raya karena barang tersebut tidak sesuai dengan spesifikasi yang dipesan. Untuk mencatat transaksi pengembalian barang dari pelanggan ini di Accurate Online, menu yang digunakan adalah ...",
        "A": "Menu Pembelian → Retur Pembelian",
        "B": "Menu Penjualan → Retur Penjualan",
        "C": "Menu Persediaan → Penyesuaian Stok",
        "D": "Menu Kas & Bank → Pengeluaran Kas",
        "E": "Menu Buku Besar → Jurnal Penyesuaian",
        "jawaban": "B",
    },
    {
        "no": 10,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Pembelian",
        "cp": "Peserta didik mampu mencatat transaksi retur pembelian di Accurate Online",
        "ipk": "Mencatat dokumen retur pembelian kepada pemasok di Accurate Online",
        "indikator": "Diberikan skenario pengembalian barang kepada pemasok, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "PT Sumber Makmur menerima kiriman barang dari pemasoknya, namun setelah diperiksa sebanyak 10 unit barang ditemukan dalam kondisi rusak dan tidak layak digunakan. PT Sumber Makmur memutuskan untuk mengembalikan barang tersebut kepada pemasok. Di Accurate Online, transaksi retur pembelian ini dicatat melalui ...",
        "A": "Menu Penjualan → Retur Penjualan",
        "B": "Menu Pembelian → Retur Pembelian",
        "C": "Menu Persediaan → Transfer Stok",
        "D": "Menu Kas & Bank → Penerimaan Kas",
        "E": "Menu Buku Besar → Jurnal Umum",
        "jawaban": "B",
    },
    {
        "no": 11,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Saldo Awal Persediaan Barang",
        "cp": "Peserta didik mampu menginput saldo awal persediaan barang dagangan di Accurate Online",
        "ipk": "Menginput saldo awal stok barang menggunakan menu penyesuaian stok di Accurate Online",
        "indikator": "Diberikan kasus perusahaan yang baru menggunakan Accurate Online dan memiliki stok barang, peserta didik dapat menentukan menu input saldo awal stok",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Sebuah perusahaan dagang yang baru pertama kali menggunakan Accurate Online memiliki stok barang senilai Rp50.000.000 yang perlu dimasukkan ke dalam sistem. Input saldo awal stok barang dagangan ini di Accurate Online dilakukan melalui menu ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Pembelian → Faktur Pembelian",
        "C": "Menu Persediaan → Penyesuaian Stok (Stock Adjustment)",
        "D": "Menu Buku Besar → Saldo Awal",
        "E": "Menu Laporan → Laporan Persediaan",
        "jawaban": "C",
    },
    {
        "no": 12,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Penerimaan Pembayaran Piutang",
        "cp": "Peserta didik mampu mencatat penerimaan pembayaran dari pelanggan di Accurate Online",
        "ipk": "Mencatat penerimaan pembayaran piutang dari pelanggan menggunakan menu Customer Receipt",
        "indikator": "Diberikan data pembayaran piutang dari pelanggan, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Pelanggan PT Harapan Jaya melakukan pembayaran atas piutangnya kepada PT Maju Bersama sebesar Rp8.000.000 melalui transfer bank. Akuntan PT Maju Bersama harus mencatat penerimaan pembayaran ini di Accurate Online. Pencatatan penerimaan pembayaran piutang dari pelanggan dilakukan melalui menu ...",
        "A": "Menu Pembelian → Pembayaran Pemasok",
        "B": "Menu Penjualan → Penerimaan Pelanggan (Customer Receipt)",
        "C": "Menu Persediaan → Transfer Stok",
        "D": "Menu Kas & Bank → Transfer Kas",
        "E": "Menu Buku Besar → Jurnal Umum",
        "jawaban": "B",
    },
    {
        "no": 13,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pembayaran Utang kepada Pemasok",
        "cp": "Peserta didik mampu mencatat pembayaran utang kepada pemasok di Accurate Online",
        "ipk": "Mencatat pembayaran utang usaha kepada pemasok menggunakan menu Vendor Payment",
        "indikator": "Diberikan data pembayaran utang kepada pemasok, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "PT Karya Mandiri melakukan pembayaran atas utang kepada PT Supplier Utama sebesar Rp12.000.000 menggunakan cek giro. Transaksi pembayaran utang dagang kepada pemasok ini di Accurate Online dicatat melalui menu ...",
        "A": "Menu Penjualan → Penerimaan Pelanggan",
        "B": "Menu Pembelian → Pembayaran Pemasok (Vendor Payment)",
        "C": "Menu Kas & Bank → Penerimaan Kas",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Buku Besar → Saldo Awal",
        "jawaban": "B",
    },
    {
        "no": 14,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Jurnal Umum",
        "cp": "Peserta didik mampu mencatat transaksi menggunakan jurnal umum di Accurate Online",
        "ipk": "Menggunakan menu Jurnal Umum untuk mencatat transaksi yang tidak tersedia di menu khusus",
        "indikator": "Diberikan jenis transaksi yang memerlukan pencatatan manual, peserta didik dapat menentukan menu Jurnal Umum di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Seorang akuntan perlu mencatat transaksi penyusutan aset tetap dan koreksi saldo akun yang tidak bisa dilakukan melalui menu transaksi khusus di Accurate Online. Untuk mencatat transaksi-transaksi seperti ini secara manual menggunakan jurnal, menu yang digunakan adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Buku Besar → Jurnal Umum (General Journal)",
        "C": "Menu Pembelian → Faktur Pembelian",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Kas & Bank → Transfer Bank",
        "jawaban": "B",
    },
    {
        "no": 15,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Data Master Pemasok",
        "cp": "Peserta didik mampu membuat dan mengelola data master pemasok di Accurate Online",
        "ipk": "Membuat data master pemasok (vendor) di Accurate Online sebelum mencatat transaksi pembelian",
        "indikator": "Diberikan skenario transaksi dengan pemasok baru, peserta didik dapat mengidentifikasi menu yang tepat untuk input data pemasok",
        "tingkat": "Rendah",
        "nilai": "5",
        "soal": "PT Sejahtera Abadi akan melakukan pembelian perdana dari pemasok baru bernama CV Bahan Baku Prima. Sebelum transaksi pembelian dapat dicatat, data master pemasok tersebut harus dibuat terlebih dahulu. Di Accurate Online, untuk menginput data master pemasok (vendor/supplier) dilakukan melalui ...",
        "A": "Menu Penjualan → Data Pelanggan",
        "B": "Menu Pembelian → Data Pemasok (Vendor List)",
        "C": "Menu Persediaan → Data Barang",
        "D": "Menu Buku Besar → Daftar Akun",
        "E": "Menu Laporan → Laporan Pembelian",
        "jawaban": "B",
    },
    {
        "no": 16,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Metode Pencatatan Persediaan",
        "cp": "Peserta didik mampu menjelaskan metode pencatatan persediaan yang tersedia di Accurate Online",
        "ipk": "Membedakan metode FIFO, LIFO, dan Average dalam pencatatan persediaan",
        "indikator": "Diberikan deskripsi metode pencatatan persediaan, peserta didik dapat mengidentifikasi nama metode yang dimaksud",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Accurate Online mendukung beberapa metode pencatatan persediaan yang dapat dipilih sesuai kebijakan perusahaan. Salah satu metode menyatakan bahwa barang yang pertama kali masuk ke gudang adalah barang yang pertama kali harus dikeluarkan saat terjadi penjualan. Metode pencatatan persediaan tersebut adalah ...",
        "A": "LIFO (Last In First Out)",
        "B": "Average (Rata-rata Bergerak)",
        "C": "FIFO (First In First Out)",
        "D": "Specific Identification (Identifikasi Khusus)",
        "E": "Lower of Cost or Market",
        "jawaban": "C",
    },
    {
        "no": 17,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Penyesuaian Stok (Stock Adjustment)",
        "cp": "Peserta didik mampu melakukan penyesuaian stok barang di Accurate Online berdasarkan hasil stock opname",
        "ipk": "Melakukan stock adjustment di Accurate Online setelah ditemukan selisih stok dari stock opname",
        "indikator": "Diberikan hasil stock opname yang menunjukkan selisih, peserta didik dapat menentukan menu penyesuaian stok",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Setelah dilakukan stock opname, ditemukan bahwa stok fisik barang jenis A hanya 85 unit, sementara catatan di Accurate Online menunjukkan 90 unit. Terdapat selisih kurang sebanyak 5 unit. Untuk menyesuaikan data stok di Accurate Online agar sesuai dengan kondisi fisik, fitur yang digunakan adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Pembelian → Purchase Order",
        "C": "Menu Persediaan → Penyesuaian Stok (Stock Adjustment)",
        "D": "Menu Buku Besar → Jurnal Umum",
        "E": "Menu Kas & Bank → Transfer Kas",
        "jawaban": "C",
    },
    {
        "no": 18,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Laporan Buku Besar",
        "cp": "Peserta didik mampu mengakses dan membaca laporan Buku Besar di Accurate Online",
        "ipk": "Mengidentifikasi jenis laporan yang menampilkan rincian mutasi setiap akun",
        "indikator": "Diberikan kebutuhan informasi mutasi akun, peserta didik dapat menentukan jenis laporan yang tepat di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Manajer keuangan PT Abadi Sejahtera meminta akuntan untuk menyajikan rincian setiap transaksi yang mempengaruhi akun Piutang Usaha selama bulan April 2025, termasuk tanggal, keterangan, dan jumlah setiap transaksi. Laporan yang harus dihasilkan dari Accurate Online adalah ...",
        "A": "Laporan Neraca (Balance Sheet)",
        "B": "Laporan Laba Rugi (Income Statement)",
        "C": "Laporan Buku Besar (General Ledger)",
        "D": "Laporan Arus Kas (Cash Flow Statement)",
        "E": "Laporan Neraca Saldo (Trial Balance)",
        "jawaban": "C",
    },
    {
        "no": 19,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Neraca Saldo (Trial Balance)",
        "cp": "Peserta didik mampu menghasilkan dan menginterpretasikan laporan Neraca Saldo dari Accurate Online",
        "ipk": "Mengakses laporan Neraca Saldo (Trial Balance) di Accurate Online melalui menu Laporan",
        "indikator": "Diberikan kebutuhan laporan neraca saldo, peserta didik dapat menentukan langkah yang tepat untuk mengaksesnya",
        "tingkat": "Sedang",
        "nilai": "5",
        "soal": "Pada akhir periode akuntansi, akuntan PT Maju Bersama perlu menyusun laporan yang memuat daftar semua akun beserta saldo debit dan kredit masing-masing untuk memverifikasi keseimbangan pencatatan. Di Accurate Online, laporan Neraca Saldo (Trial Balance) dapat diakses melalui ...",
        "A": "Menu Penjualan → Laporan Penjualan",
        "B": "Menu Pembelian → Laporan Pembelian",
        "C": "Menu Laporan → Buku Besar → Neraca Saldo",
        "D": "Menu Persediaan → Laporan Persediaan",
        "E": "Menu Kas & Bank → Laporan Kas",
        "jawaban": "C",
    },
    {
        "no": 20,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Transaksi Kas & Bank",
        "cp": "Peserta didik mampu mencatat transaksi kas masuk dan keluar yang tidak terkait penjualan/pembelian",
        "ipk": "Mencatat pengeluaran kas yang tidak terkait langsung dengan pembelian barang dagangan",
        "indikator": "Diberikan transaksi pengeluaran biaya operasional, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sulit",
        "nilai": "5",
        "soal": "PT Makmur Jaya membayar biaya sewa gedung kantor sebesar Rp3.000.000 secara tunai. Transaksi ini bukan merupakan pembelian barang dagangan, melainkan pengeluaran biaya operasional. Di Accurate Online, pencatatan transaksi pengeluaran biaya operasional seperti ini dilakukan melalui ...",
        "A": "Menu Pembelian → Faktur Pembelian",
        "B": "Menu Penjualan → Faktur Penjualan",
        "C": "Menu Kas & Bank → Pengeluaran Kas/Bank",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Buku Besar → Saldo Awal",
        "jawaban": "C",
    },
]

# 10 Soal Essay
essay_data = [
    {
        "no": 1,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Setup Data Awal Perusahaan",
        "cp": "Peserta didik mampu melakukan setup data awal perusahaan di Accurate Online secara lengkap dan benar",
        "ipk": "Menjelaskan prosedur pembuatan database dan setup awal perusahaan di Accurate Online",
        "indikator": "Diberikan kasus perusahaan baru yang akan menggunakan Accurate Online, peserta didik dapat menjelaskan langkah-langkah setup awal",
        "tingkat": "Sedang",
        "nilai": "10",
        "narasi": (
            "Accurate Online merupakan software akuntansi berbasis cloud (cloud-based accounting software) "
            "yang dikembangkan oleh PT Cipta Piranti Sejahtera. Software ini telah banyak digunakan oleh "
            "perusahaan-perusahaan di Indonesia karena kemudahannya dalam mencatat dan mengelola data "
            "keuangan secara terintegrasi.\n\n"
            "Sebelum mulai menggunakan Accurate Online, pengguna harus melakukan setup atau pengaturan "
            "awal perusahaan terlebih dahulu. Proses setup ini meliputi pengisian berbagai informasi dasar "
            "perusahaan yang akan menjadi landasan dari semua proses pencatatan akuntansi selanjutnya. "
            "Kesalahan dalam proses setup awal dapat berdampak pada keseluruhan laporan keuangan yang "
            "dihasilkan.\n\n"
            "Bayangkan kamu adalah seorang staf akuntansi di perusahaan yang baru saja memutuskan untuk "
            "menggunakan Accurate Online sebagai software akuntansinya."
        ),
        "pertanyaan": "Jelaskan minimal 5 langkah yang harus dilakukan dalam proses setup data awal perusahaan di Accurate Online beserta penjelasan singkat setiap langkahnya!",
        "jawaban": (
            "Langkah-langkah setup data awal perusahaan di Accurate Online:\n\n"
            "1. Login ke Accurate Online melalui website resmi (app.accurate.id) menggunakan akun yang telah terdaftar.\n\n"
            "2. Membuat database perusahaan baru (New Company) dengan mengklik tombol Buat Perusahaan Baru.\n\n"
            "3. Mengisi informasi umum perusahaan: nama perusahaan, alamat lengkap, nomor telepon, email, NPWP, dan logo perusahaan.\n\n"
            "4. Menentukan periode fiskal (tahun buku) perusahaan, misalnya 1 Januari 2025 s.d. 31 Desember 2025.\n\n"
            "5. Menentukan mata uang fungsional yang digunakan (misalnya Rupiah/IDR).\n\n"
            "6. Memilih metode pencatatan persediaan: FIFO (First In First Out) atau Average (Rata-rata).\n\n"
            "7. Mengatur Daftar Akun (Chart of Accounts) sesuai dengan kebutuhan perusahaan, termasuk menambah atau menonaktifkan akun yang tidak diperlukan.\n\n"
            "8. Memasukkan saldo awal akun (Opening Balance) berdasarkan neraca perusahaan sebelumnya."
        ),
    },
    {
        "no": 2,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Daftar Akun (Chart of Accounts)",
        "cp": "Peserta didik mampu membuat dan mengelola daftar akun di Accurate Online sesuai kebutuhan perusahaan",
        "ipk": "Menjelaskan pengelompokan akun dalam Accurate Online beserta contoh masing-masing",
        "indikator": "Diberikan informasi tentang sistem pengkodean akun, peserta didik dapat menjelaskan kelompok akun beserta contohnya",
        "tingkat": "Sedang",
        "nilai": "10",
        "narasi": (
            "Daftar Akun (Chart of Accounts) merupakan komponen fundamental dalam sistem akuntansi "
            "berbantuan komputer seperti Accurate Online. Daftar akun berisi seluruh akun yang akan "
            "digunakan perusahaan untuk mencatat semua transaksi keuangannya. Setiap akun memiliki kode "
            "unik, nama, dan tipe yang berbeda-beda.\n\n"
            "Accurate Online telah menyediakan daftar akun standar secara default yang dapat dimodifikasi "
            "sesuai kebutuhan spesifik perusahaan. Sistem pengkodean akun menggunakan angka sebagai "
            "identifikasi kelompok akun. Pemahaman yang mendalam tentang pengelompokan akun sangat "
            "penting karena berpengaruh langsung terhadap keakuratan laporan keuangan yang dihasilkan.\n\n"
            "Sebagai seorang pelajar Komputer Akuntansi Kelas XI, kamu perlu memahami struktur daftar akun "
            "di Accurate Online agar dapat menggunakannya dengan benar dalam mencatat transaksi."
        ),
        "pertanyaan": "Jelaskan pengelompokan akun dalam Accurate Online berdasarkan kode akun (kelompok 1 sampai dengan kelompok 5), beserta minimal 3 contoh akun untuk setiap kelompok!",
        "jawaban": (
            "Pengelompokan akun dalam Accurate Online:\n\n"
            "1. Kelompok 1 – Aset (Harta): Akun-akun yang mencerminkan kekayaan perusahaan.\n"
            "   Contoh: Kas (1-10100), Bank BCA (1-10200), Piutang Usaha (1-10300), Persediaan Barang (1-10500), Peralatan (1-15000).\n\n"
            "2. Kelompok 2 – Liabilitas (Kewajiban): Akun-akun yang mencerminkan kewajiban/utang perusahaan.\n"
            "   Contoh: Utang Usaha (2-10100), Utang Bank (2-11000), Utang Gaji (2-10500).\n\n"
            "3. Kelompok 3 – Ekuitas (Modal): Akun-akun yang mencerminkan hak pemilik atas perusahaan.\n"
            "   Contoh: Modal Pemilik (3-10000), Laba Ditahan (3-20000), Prive Pemilik (3-30000).\n\n"
            "4. Kelompok 4 – Pendapatan: Akun-akun yang mencatat penghasilan perusahaan.\n"
            "   Contoh: Pendapatan Penjualan (4-10000), Pendapatan Jasa (4-20000), Diskon Penjualan (4-30000).\n\n"
            "5. Kelompok 5 – Beban: Akun-akun yang mencatat pengeluaran/biaya perusahaan.\n"
            "   Contoh: Harga Pokok Penjualan (5-10000), Beban Gaji (5-20000), Beban Sewa (5-30000)."
        ),
    },
    {
        "no": 3,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Input Saldo Awal Akun",
        "cp": "Peserta didik mampu menginput saldo awal seluruh akun neraca di Accurate Online untuk perusahaan yang sudah berjalan",
        "ipk": "Menjelaskan prosedur input saldo awal dan alasan mengapa saldo debit harus sama dengan saldo kredit",
        "indikator": "Diberikan kondisi perusahaan yang baru menggunakan Accurate Online, peserta didik dapat menjelaskan prosedur dan prinsip input saldo awal",
        "tingkat": "Sedang",
        "nilai": "10",
        "narasi": (
            "Ketika sebuah perusahaan yang sudah berjalan mulai menggunakan Accurate Online, salah satu "
            "tahapan paling krusial adalah memasukkan saldo awal dari semua akun yang ada. Saldo awal "
            "ini berasal dari neraca saldo perusahaan pada akhir periode sebelumnya.\n\n"
            "Tanpa saldo awal yang benar dan lengkap, laporan keuangan yang dihasilkan oleh Accurate "
            "Online tidak akan mencerminkan kondisi keuangan perusahaan yang sesungguhnya. Proses input "
            "saldo awal harus dilakukan dengan sangat teliti dan cermat karena kesalahan sekecil apapun "
            "dapat mengakibatkan laporan keuangan menjadi tidak akurat.\n\n"
            "Prinsip dasar akuntansi menyatakan bahwa dalam setiap pencatatan transaksi, total saldo debit "
            "harus selalu sama dengan total saldo kredit. Hal ini berlaku pula dalam input saldo awal di "
            "Accurate Online."
        ),
        "pertanyaan": "Jelaskan apa yang dimaksud dengan saldo awal dalam Accurate Online, uraikan prosedur input saldo awal akun secara lengkap, dan mengapa total saldo debit harus sama dengan total saldo kredit!",
        "jawaban": (
            "Pengertian Saldo Awal:\n"
            "Saldo awal (opening balance) adalah saldo dari setiap akun pada saat pertama kali perusahaan "
            "mulai menggunakan Accurate Online, yang bersumber dari laporan neraca atau neraca saldo akhir "
            "periode sebelumnya.\n\n"
            "Prosedur Input Saldo Awal di Accurate Online:\n"
            "1. Buka Menu Buku Besar → pilih Saldo Awal\n"
            "2. Tentukan tanggal saldo awal (biasanya tanggal awal periode penggunaan Accurate Online)\n"
            "3. Masukkan saldo masing-masing akun sesuai kolom debit atau kredit:\n"
            "   - Akun Aset: masukkan di kolom Debit\n"
            "   - Akun Liabilitas dan Ekuitas: masukkan di kolom Kredit\n"
            "4. Periksa total saldo – pastikan Total Debit = Total Kredit\n"
            "5. Klik Simpan\n\n"
            "Alasan Debit = Kredit:\n"
            "Prinsip ini merupakan dasar dari sistem akuntansi double-entry bookkeeping (pencatatan ganda). "
            "Setiap transaksi selalu mempengaruhi minimal dua akun dengan jumlah yang sama namun di sisi "
            "berlawanan (debit dan kredit), sehingga persamaan akuntansi Aset = Liabilitas + Ekuitas "
            "selalu terjaga keseimbangannya."
        ),
    },
    {
        "no": 4,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pencatatan Transaksi Penjualan Kredit",
        "cp": "Peserta didik mampu mencatat siklus lengkap transaksi penjualan kredit di Accurate Online",
        "ipk": "Menjelaskan prosedur pencatatan penjualan kredit dari pembuatan faktur hingga pelunasan piutang",
        "indikator": "Diberikan skenario transaksi penjualan kredit, peserta didik dapat menjelaskan prosedur lengkap pencatatannya di Accurate Online",
        "tingkat": "Sedang",
        "nilai": "10",
        "narasi": (
            "Transaksi penjualan merupakan salah satu transaksi utama yang paling sering dilakukan oleh "
            "perusahaan dagang. Di Accurate Online, terdapat alur yang sistematis dalam proses pencatatan "
            "penjualan, mulai dari pembuatan sales quotation (penawaran harga), sales order (pesanan "
            "penjualan), hingga faktur penjualan (invoice) yang merupakan dokumen resmi tagihan kepada "
            "pelanggan.\n\n"
            "Faktur penjualan yang dibuat di Accurate Online akan secara otomatis menghasilkan jurnal "
            "akuntansi yang sesuai. Sistem juga akan secara otomatis memperbarui saldo piutang usaha dan "
            "mengurangi stok persediaan barang yang terjual.\n\n"
            "Siklus penjualan kredit belum selesai sampai pelanggan melunasi pembayarannya. Pelunasan "
            "piutang harus pula dicatat agar saldo piutang usaha menjadi akurat."
        ),
        "pertanyaan": "Jelaskan prosedur lengkap pencatatan transaksi penjualan kredit di Accurate Online, mulai dari pembuatan faktur penjualan hingga pelunasan piutang oleh pelanggan! Sertakan jurnal yang terbentuk pada setiap tahap.",
        "jawaban": (
            "Prosedur Penjualan Kredit di Accurate Online:\n\n"
            "TAHAP 1 – Pembuatan Faktur Penjualan:\n"
            "1. Buka Menu Penjualan → Faktur Penjualan → Baru\n"
            "2. Pilih nama pelanggan dari daftar pelanggan yang sudah ada\n"
            "3. Isi tanggal transaksi dan nomor faktur\n"
            "4. Pilih barang yang dijual, isi kuantitas dan harga jual\n"
            "5. Tentukan syarat pembayaran (contoh: Net 30)\n"
            "6. Klik Simpan\n"
            "Jurnal otomatis: Debit Piutang Usaha, Kredit Penjualan\n"
            "                Debit HPP, Kredit Persediaan Barang\n\n"
            "TAHAP 2 – Penerimaan Pembayaran dari Pelanggan:\n"
            "1. Buka Menu Penjualan → Penerimaan Pelanggan (Customer Receipt) → Baru\n"
            "2. Pilih nama pelanggan\n"
            "3. Pilih faktur yang dilunasi dan isi jumlah yang diterima\n"
            "4. Pilih akun kas/bank yang menerima pembayaran\n"
            "5. Isi tanggal penerimaan dan klik Simpan\n"
            "Jurnal otomatis: Debit Kas/Bank, Kredit Piutang Usaha"
        ),
    },
    {
        "no": 5,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Penjualan",
        "cp": "Peserta didik mampu mencatat transaksi retur penjualan di Accurate Online dengan benar",
        "ipk": "Menjelaskan prosedur pencatatan retur penjualan dan jurnal yang terbentuk",
        "indikator": "Diberikan skenario pengembalian barang dari pelanggan, peserta didik dapat menjelaskan prosedur dan jurnal retur penjualan",
        "tingkat": "Sedang",
        "nilai": "10",
        "narasi": (
            "Retur penjualan terjadi ketika pelanggan mengembalikan barang yang telah dibeli karena berbagai "
            "alasan, antara lain: barang rusak/cacat saat diterima, barang tidak sesuai spesifikasi yang "
            "dipesan, kelebihan pengiriman barang, atau barang kadaluarsa.\n\n"
            "Di Accurate Online, proses retur penjualan terintegrasi secara langsung dengan faktur "
            "penjualan yang menjadi referensinya. Integrasi ini memudahkan pelacakan dan memastikan "
            "konsistensi data antara penjualan dan retuurnya.\n\n"
            "Pencatatan retur penjualan yang tepat dan tepat waktu sangat penting untuk memastikan bahwa "
            "saldo piutang usaha, saldo persediaan barang, dan laporan penjualan mencerminkan kondisi yang "
            "sebenarnya. Keterlambatan atau kesalahan dalam mencatat retur dapat menyebabkan laporan "
            "keuangan menjadi tidak akurat."
        ),
        "pertanyaan": "Jelaskan prosedur pencatatan retur penjualan di Accurate Online secara lengkap, dan sebutkan jurnal akuntansi yang secara otomatis terbentuk dari transaksi retur penjualan tersebut!",
        "jawaban": (
            "Prosedur Retur Penjualan di Accurate Online:\n\n"
            "1. Buka Menu Penjualan → Retur Penjualan → Baru\n"
            "2. Pilih nama pelanggan yang melakukan pengembalian barang\n"
            "3. Di kolom referensi, pilih nomor faktur penjualan asli yang menjadi dasar retur\n"
            "4. Sistem otomatis menampilkan daftar barang dari faktur tersebut\n"
            "5. Pilih barang yang diretur dan masukkan kuantitas yang dikembalikan\n"
            "6. Isi keterangan/alasan retur (opsional)\n"
            "7. Tentukan tanggal retur\n"
            "8. Klik Simpan\n\n"
            "Jurnal yang Terbentuk Otomatis:\n"
            "Untuk retur penjualan kredit:\n"
            "- Debit: Retur & Potongan Penjualan (mengurangi pendapatan)\n"
            "- Kredit: Piutang Usaha (mengurangi tagihan kepada pelanggan)\n\n"
            "Untuk pengembalian stok barang:\n"
            "- Debit: Persediaan Barang (barang kembali ke gudang)\n"
            "- Kredit: Harga Pokok Penjualan / HPP (mengurangi HPP)"
        ),
    },
    {
        "no": 6,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pencatatan Transaksi Pembelian Kredit",
        "cp": "Peserta didik mampu mencatat siklus lengkap transaksi pembelian kredit di Accurate Online",
        "ipk": "Menjelaskan prosedur pencatatan pembelian kredit dari Purchase Order hingga pembayaran utang",
        "indikator": "Diberikan skenario pembelian kredit, peserta didik dapat menjelaskan prosedur dan jurnal lengkap pembelian kredit di Accurate Online",
        "tingkat": "Sulit",
        "nilai": "10",
        "narasi": (
            "Transaksi pembelian barang dari pemasok merupakan kegiatan rutin yang dilakukan oleh perusahaan "
            "dagang. Di Accurate Online, proses pembelian dapat dimulai dengan pembuatan Purchase Order "
            "(PO) sebagai surat pesanan kepada pemasok, dilanjutkan dengan penerimaan barang, dan "
            "diakhiri dengan pembuatan faktur pembelian berdasarkan tagihan dari pemasok.\n\n"
            "Pencatatan pembelian yang akurat dan tepat waktu sangat penting untuk memantau posisi utang "
            "kepada pemasok dan memastikan ketersediaan stok barang yang memadai. Setiap transaksi "
            "pembelian yang dicatat di Accurate Online akan secara otomatis memperbarui saldo utang usaha "
            "dan menambah jumlah persediaan barang.\n\n"
            "Siklus pembelian kredit belum lengkap sampai utang kepada pemasok dilunasi. Pembayaran utang "
            "juga harus dicatat dengan benar agar posisi kas dan saldo utang selalu akurat."
        ),
        "pertanyaan": "Jelaskan prosedur lengkap pencatatan transaksi pembelian kredit di Accurate Online, mulai dari Purchase Order (PO) hingga pembayaran utang kepada pemasok! Sertakan jurnal yang terbentuk pada setiap tahap.",
        "jawaban": (
            "Prosedur Pembelian Kredit di Accurate Online:\n\n"
            "TAHAP 1 – Pembuatan Purchase Order (PO):\n"
            "1. Buka Menu Pembelian → Purchase Order → Baru\n"
            "2. Pilih nama pemasok, isi tanggal PO\n"
            "3. Pilih barang yang dipesan, isi kuantitas dan harga\n"
            "4. Klik Simpan\n"
            "(PO tidak menghasilkan jurnal, hanya sebagai dokumen pemesanan)\n\n"
            "TAHAP 2 – Pembuatan Faktur Pembelian:\n"
            "1. Buka Menu Pembelian → Faktur Pembelian → Baru\n"
            "2. Pilih pemasok dan referensikan ke PO yang sudah dibuat\n"
            "3. Isi nomor faktur dari pemasok dan tanggal jatuh tempo\n"
            "4. Klik Simpan\n"
            "Jurnal otomatis: Debit Persediaan Barang, Kredit Utang Usaha\n\n"
            "TAHAP 3 – Pembayaran Utang kepada Pemasok:\n"
            "1. Buka Menu Pembelian → Pembayaran Pemasok (Vendor Payment) → Baru\n"
            "2. Pilih nama pemasok\n"
            "3. Pilih faktur yang akan dibayar dan masukkan jumlah pembayaran\n"
            "4. Pilih akun kas/bank yang digunakan untuk membayar\n"
            "5. Klik Simpan\n"
            "Jurnal otomatis: Debit Utang Usaha, Kredit Kas/Bank"
        ),
    },
    {
        "no": 7,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Pembelian",
        "cp": "Peserta didik mampu mencatat transaksi retur pembelian di Accurate Online dengan benar",
        "ipk": "Menjelaskan prosedur dan jurnal retur pembelian beserta contoh kasusnya",
        "indikator": "Diberikan contoh kasus retur pembelian, peserta didik dapat menjelaskan prosedur pencatatan dan jurnalnya",
        "tingkat": "Sulit",
        "nilai": "10",
        "narasi": (
            "Retur pembelian terjadi ketika perusahaan mengembalikan sebagian atau seluruh barang yang "
            "telah dibeli dari pemasok karena berbagai alasan, seperti: barang cacat/rusak, barang tidak "
            "sesuai dengan spesifikasi dalam Purchase Order, kelebihan pengiriman dari jumlah yang dipesan, "
            "atau barang sudah kadaluarsa saat diterima.\n\n"
            "Di Accurate Online, fitur retur pembelian dirancang untuk terintegrasi langsung dengan faktur "
            "pembelian yang menjadi referensinya. Dengan demikian, sistem dapat secara otomatis "
            "memperbarui saldo utang kepada pemasok dan mengurangi jumlah stok barang yang dikembalikan.\n\n"
            "Contoh Kasus: CV Mitra Abadi membeli 100 unit barang dari PT Jaya Supplier seharga "
            "Rp5.000.000 secara kredit. Setelah dilakukan pemeriksaan, ditemukan 10 unit barang dalam "
            "kondisi rusak sehingga CV Mitra Abadi memutuskan mengembalikannya kepada PT Jaya Supplier."
        ),
        "pertanyaan": "Berdasarkan contoh kasus di atas, jelaskan prosedur pencatatan retur pembelian di Accurate Online secara lengkap dan sebutkan jurnal yang terbentuk dari transaksi tersebut! (Harga per unit = Rp50.000)",
        "jawaban": (
            "Prosedur Retur Pembelian di Accurate Online (berdasarkan kasus CV Mitra Abadi):\n\n"
            "1. Buka Menu Pembelian → Retur Pembelian → Baru\n"
            "2. Pilih nama pemasok: PT Jaya Supplier\n"
            "3. Di kolom referensi, pilih nomor faktur pembelian asli\n"
            "4. Sistem otomatis menampilkan daftar barang yang dibeli\n"
            "5. Pilih barang yang diretur dan masukkan kuantitas = 10 unit\n"
            "6. Sistem otomatis menghitung nilai retur = 10 unit × Rp50.000 = Rp500.000\n"
            "7. Isi keterangan alasan retur: 'Barang rusak'\n"
            "8. Tentukan tanggal retur pembelian\n"
            "9. Klik Simpan\n\n"
            "Jurnal yang Terbentuk Otomatis:\n"
            "- Debit: Utang Usaha (PT Jaya Supplier)   Rp500.000\n"
            "         (mengurangi utang karena barang dikembalikan)\n"
            "- Kredit: Persediaan Barang               Rp500.000\n"
            "         (mengurangi stok karena barang keluar kembali ke pemasok)\n\n"
            "Dengan pencatatan ini, utang CV Mitra Abadi kepada PT Jaya Supplier berkurang dari "
            "Rp5.000.000 menjadi Rp4.500.000, dan stok barang berkurang 10 unit."
        ),
    },
    {
        "no": 8,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Penyesuaian Stok (Stock Adjustment)",
        "cp": "Peserta didik mampu melakukan penyesuaian stok berdasarkan hasil stock opname di Accurate Online",
        "ipk": "Menjelaskan prosedur stock opname dan stock adjustment di Accurate Online",
        "indikator": "Diberikan hasil stock opname dengan selisih stok, peserta didik dapat menjelaskan prosedur penyesuaian stok dan jurnal yang terbentuk",
        "tingkat": "Sulit",
        "nilai": "10",
        "narasi": (
            "Stock opname adalah kegiatan penghitungan fisik stok barang yang ada di gudang secara "
            "langsung, kemudian membandingkan hasilnya dengan data stok yang tercatat di sistem akuntansi "
            "(Accurate Online). Stock opname biasanya dilakukan secara berkala, misalnya setiap akhir "
            "bulan atau akhir tahun.\n\n"
            "Perbedaan antara stok fisik dan stok di sistem dapat terjadi karena berbagai faktor, antara "
            "lain: barang hilang, barang rusak yang belum dilaporkan, kesalahan pencatatan, atau barang "
            "yang tercecer. Jika ditemukan selisih, maka harus dilakukan penyesuaian stok (stock "
            "adjustment) di Accurate Online agar data sistem kembali mencerminkan kondisi fisik yang "
            "sebenarnya.\n\n"
            "Contoh Kasus: Hasil stock opname menunjukkan bahwa stok fisik Barang X hanya 85 unit, "
            "sedangkan catatan di Accurate Online menunjukkan 90 unit. Terdapat selisih kurang 5 unit "
            "dengan harga pokok Rp20.000 per unit."
        ),
        "pertanyaan": "Jelaskan apa yang dimaksud dengan stock opname, uraikan prosedur penyesuaian stok (stock adjustment) di Accurate Online berdasarkan kasus di atas, dan sebutkan jurnal yang terbentuk!",
        "jawaban": (
            "Pengertian Stock Opname:\n"
            "Stock opname adalah proses penghitungan fisik seluruh persediaan barang di gudang dan "
            "membandingkannya dengan catatan di sistem akuntansi untuk memastikan keakuratan data stok.\n\n"
            "Prosedur Stock Adjustment di Accurate Online (kasus Barang X):\n"
            "1. Lakukan penghitungan fisik barang di gudang → Barang X: 85 unit\n"
            "2. Bandingkan dengan stok di Accurate Online → tercatat 90 unit\n"
            "3. Selisih = 90 – 85 = 5 unit (stok fisik lebih sedikit)\n"
            "4. Buka Menu Persediaan → Penyesuaian Stok (Stock Adjustment) → Baru\n"
            "5. Pilih barang: Barang X\n"
            "6. Masukkan kuantitas aktual: 85 unit\n"
            "7. Sistem otomatis menghitung selisih: -5 unit\n"
            "8. Masukkan keterangan: 'Penyesuaian hasil stock opname'\n"
            "9. Klik Simpan\n\n"
            "Jurnal yang Terbentuk:\n"
            "Nilai selisih = 5 unit × Rp20.000 = Rp100.000\n"
            "- Debit: Beban Kerugian Selisih Persediaan   Rp100.000\n"
            "- Kredit: Persediaan Barang                  Rp100.000\n"
            "(mencatat pengurangan stok akibat selisih hasil stock opname)"
        ),
    },
    {
        "no": 9,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Neraca Saldo (Trial Balance)",
        "cp": "Peserta didik mampu menghasilkan dan menginterpretasikan laporan Neraca Saldo dari Accurate Online",
        "ipk": "Menjelaskan pengertian, fungsi, dan cara mengakses Neraca Saldo di Accurate Online",
        "indikator": "Diberikan kebutuhan laporan neraca saldo, peserta didik dapat menjelaskan pengertian, fungsi, dan cara mengaksesnya",
        "tingkat": "Sedang",
        "nilai": "10",
        "narasi": (
            "Neraca Saldo (Trial Balance) merupakan salah satu laporan penting yang dihasilkan dalam "
            "proses akuntansi. Laporan ini memuat daftar semua akun beserta saldo debit atau kredit "
            "masing-masing akun pada suatu periode tertentu.\n\n"
            "Neraca saldo menjadi alat verifikasi pertama sebelum laporan keuangan disusun. Jika total "
            "saldo debit sama dengan total saldo kredit, maka secara matematis pencatatan transaksi telah "
            "seimbang. Namun perlu diingat bahwa neraca saldo yang seimbang tidak menjamin bebas dari "
            "semua jenis kesalahan pencatatan.\n\n"
            "Di Accurate Online, neraca saldo dihasilkan secara otomatis berdasarkan seluruh transaksi "
            "yang telah dicatat dalam sistem. Laporan ini dapat dicetak atau diekspor ke format Excel "
            "untuk keperluan analisis lebih lanjut."
        ),
        "pertanyaan": "Jelaskan pengertian Neraca Saldo (Trial Balance) dalam akuntansi, uraikan minimal 3 fungsi Neraca Saldo, dan jelaskan langkah-langkah mengakses laporan Neraca Saldo di Accurate Online!",
        "jawaban": (
            "Pengertian Neraca Saldo:\n"
            "Neraca Saldo (Trial Balance) adalah daftar yang memuat semua akun beserta saldo debit atau "
            "kredit masing-masing akun pada akhir suatu periode akuntansi, yang digunakan untuk "
            "memverifikasi keseimbangan antara total saldo debit dan total saldo kredit.\n\n"
            "Fungsi Neraca Saldo:\n"
            "1. Memverifikasi keseimbangan pencatatan: memastikan total debit = total kredit sebagai bukti "
            "   matematis kebenaran double-entry bookkeeping\n"
            "2. Dasar penyusunan laporan keuangan: menjadi titik awal dalam menyusun Laporan Laba Rugi, "
            "   Neraca, dan laporan keuangan lainnya\n"
            "3. Alat deteksi kesalahan: membantu mengidentifikasi kemungkinan adanya kesalahan posting "
            "   atau transaksi yang belum dicatat\n"
            "4. Memberikan gambaran umum posisi keuangan perusahaan secara ringkas\n\n"
            "Langkah Mengakses Neraca Saldo di Accurate Online:\n"
            "1. Klik Menu Laporan (Reports) di panel menu utama\n"
            "2. Pilih submenu Buku Besar (General Ledger)\n"
            "3. Pilih Neraca Saldo (Trial Balance)\n"
            "4. Tentukan periode laporan: isi tanggal awal dan tanggal akhir\n"
            "5. Klik tombol Tampilkan / Generate\n"
            "6. Laporan Neraca Saldo akan ditampilkan dan siap dicetak atau diekspor"
        ),
    },
    {
        "no": 10,
        "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Alur Kerja Accurate Online: Dari Setup hingga Laporan",
        "cp": "Peserta didik mampu menjelaskan alur kerja lengkap penggunaan Accurate Online dari setup awal hingga laporan keuangan",
        "ipk": "Mendeskripsikan workflow Accurate Online secara sistematis dan terstruktur",
        "indikator": "Diberikan tugas membuat deskripsi penggunaan Accurate Online, peserta didik dapat menjelaskan alur kerja lengkap dari awal hingga neraca saldo",
        "tingkat": "Sulit",
        "nilai": "10",
        "narasi": (
            "Accurate Online adalah sistem akuntansi yang dirancang untuk memudahkan proses pencatatan "
            "keuangan perusahaan secara terintegrasi. Semua modul dalam Accurate Online saling terhubung "
            "satu sama lain, sehingga data yang diinput di satu modul akan otomatis mempengaruhi modul "
            "lainnya.\n\n"
            "Agar penggunaan Accurate Online berjalan dengan optimal, pengguna harus memahami alur kerja "
            "(workflow) yang benar, mulai dari tahap persiapan awal hingga tahap pelaporan. Pemahaman "
            "tentang alur kerja ini akan membantu pengguna menghindari kesalahan input dan memastikan "
            "semua transaksi tercatat dengan tepat dan akurat.\n\n"
            "Sebagai calon tenaga akuntansi profesional, kamu dituntut untuk dapat menguasai dan "
            "menjelaskan seluruh alur penggunaan Accurate Online secara sistematis kepada pengguna lain "
            "atau kepada manajemen perusahaan."
        ),
        "pertanyaan": "Jelaskan secara sistematis dan lengkap alur kerja (workflow) penggunaan Accurate Online, mulai dari tahap setup awal perusahaan hingga dihasilkannya laporan Neraca Saldo! Sertakan setiap tahapan beserta penjelasan singkatnya.",
        "jawaban": (
            "Alur Kerja (Workflow) Accurate Online:\n\n"
            "TAHAP 1 – SETUP AWAL PERUSAHAAN\n"
            "- Buat database perusahaan baru di Accurate Online\n"
            "- Isi informasi perusahaan (nama, alamat, NPWP, periode fiskal, mata uang)\n"
            "- Pilih metode persediaan (FIFO/Average)\n\n"
            "TAHAP 2 – PENGATURAN DAFTAR AKUN\n"
            "- Review dan sesuaikan Chart of Accounts default\n"
            "- Tambah/nonaktifkan akun sesuai kebutuhan perusahaan\n\n"
            "TAHAP 3 – INPUT DATA MASTER\n"
            "- Buat daftar Pelanggan (Customer List)\n"
            "- Buat daftar Pemasok (Vendor List)\n"
            "- Buat daftar Barang/Item (Item List)\n\n"
            "TAHAP 4 – INPUT SALDO AWAL\n"
            "- Masukkan saldo awal semua akun dari neraca sebelumnya\n"
            "- Masukkan saldo awal persediaan barang\n"
            "- Masukkan saldo awal piutang dan utang\n\n"
            "TAHAP 5 – PENCATATAN TRANSAKSI HARIAN\n"
            "- Penjualan: Faktur Penjualan\n"
            "- Pembelian: Faktur Pembelian\n"
            "- Penerimaan pembayaran dari pelanggan: Customer Receipt\n"
            "- Pembayaran kepada pemasok: Vendor Payment\n"
            "- Retur penjualan/pembelian bila ada\n"
            "- Transaksi lain: Jurnal Umum / Kas & Bank\n\n"
            "TAHAP 6 – LAPORAN NERACA SALDO\n"
            "- Akses Menu Laporan → Buku Besar → Neraca Saldo\n"
            "- Tentukan periode laporan\n"
            "- Verifikasi total debit = total kredit\n"
            "- Cetak/ekspor laporan untuk analisis lebih lanjut"
        ),
    },
]

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clear_cell_content(cell):
    """Remove all paragraphs from cell except the first."""
    tc = cell._tc
    # Keep only the first paragraph, remove others
    paras = tc.findall(qn('w:p'))
    for p in paras[1:]:
        tc.remove(p)
    # Clear runs in first paragraph
    first_para = paras[0] if paras else None
    if first_para is not None:
        for r in first_para.findall(qn('w:r')):
            first_para.remove(r)
        # Remove any hyperlinks too
        for hl in first_para.findall(qn('w:hyperlink')):
            first_para.remove(hl)


def set_cell_simple(cell, text):
    """Set cell text to a simple string, preserving paragraph format."""
    clear_cell_content(cell)
    tc = cell._tc
    paras = tc.findall(qn('w:p'))
    first_para = cell.paragraphs[0]
    
    lines = text.split('\n')
    
    # Set first line in existing paragraph
    run = first_para.add_run(lines[0])
    
    # Add subsequent lines as new paragraphs
    for line in lines[1:]:
        new_para = cell.add_paragraph(line)


def update_run_text(cell, old_text, new_text):
    """Find a run containing old_text and replace it."""
    for para in cell.paragraphs:
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                return True
    return False


def fill_header_cell(cell, nama_penulis, mapel, kelas):
    """Update Nama Penulis Soal, Mata Pelajaran, Kelas in header cell."""
    for para in cell.paragraphs:
        for run in para.runs:
            if '..............di isi' in run.text:
                run.text = run.text.replace('..............di isi', nama_penulis)
            elif '.............di isi' in run.text:
                run.text = run.text.replace('.............di isi', f'{mapel} / Kelas {kelas}')
            elif '...............' in run.text:
                run.text = run.text.replace('...............', nama_penulis)


def fill_referensi_cell(cell, ref_text):
    """Fill the Buku Referensi cell."""
    clear_cell_content(cell)
    first_para = cell.paragraphs[0]
    first_para.add_run("Buku Acuan/Referensi:\n" + ref_text)


def fill_pg_table(table, q, no_soal_str=None):
    """Fill a PG kartu soal table with question data."""
    
    # Row 0 - Header: update Nama Penulis Soal dan Mata Pelajaran
    fill_header_cell(table.rows[0].cells[0], NAMA_PENULIS, MAPEL, KELAS)
    
    # Row 1-2 col 1-4: Buku Referensi
    for r in [1, 2]:
        for c in [1, 2, 3, 4]:
            try:
                cell = table.rows[r].cells[c]
                # Check if this is a unique cell (not repeated due to column merge)
                clear_cell_content(cell)
                cell.paragraphs[0].add_run(f"Buku Acuan/Referensi:\n{q['referensi'] if 'referensi' in q else REF}")
            except Exception:
                pass
    
    # Row 2 col 0: Elemen Kompetensi value
    set_cell_simple(table.rows[2].cells[0], q['elemen'])
    
    # Row 3 col 0: Materi/Sub-materi
    set_cell_simple(table.rows[3].cells[0], q['materi'])
    
    # Row 5 col 0: Capaian Pembelajaran value
    set_cell_simple(table.rows[5].cells[0], q['cp'])
    
    # Row 6 col 0: Kelas  
    set_cell_simple(table.rows[6].cells[0], f"Kelas: {KELAS}")
    
    # Row 7 col 0: Semester
    set_cell_simple(table.rows[7].cells[0], "Semester: Genap")
    
    # Row 8 col 0: Bobot nilai
    set_cell_simple(table.rows[8].cells[0], f"Bobot Soal: {q.get('nilai', '5')} poin")
    
    # Row 10 col 0: IPK
    set_cell_simple(table.rows[10].cells[0], q['ipk'])
    
    # Row 12 col 0: Indikator Soal
    set_cell_simple(table.rows[12].cells[0], q['indikator'])
    
    # Row 13 col 0: Tingkat Kesukaran
    tingkat = q.get('tingkat', 'Sedang')
    set_cell_simple(table.rows[13].cells[0], 
        f"Tingkat Kesukaran:\n{'✓ Rendah' if tingkat=='Rendah' else '  Rendah'}\n"
        f"{'✓ Sedang' if tingkat=='Sedang' else '  Sedang'}\n"
        f"{'✓ Sulit' if tingkat=='Sulit' else '  Sulit'}")
    
    # Row 5 col 2: Nomor soal
    if no_soal_str:
        set_cell_simple(table.rows[5].cells[2], no_soal_str)
    else:
        set_cell_simple(table.rows[5].cells[2], str(q['no']))
    
    # Row 8 col 2: Nilai
    set_cell_simple(table.rows[8].cells[2], f"{q.get('nilai', '5')}")
    
    # Row 11-13 col 2: Kunci Jawaban
    jawaban = q.get('jawaban', '')
    set_cell_simple(table.rows[11].cells[2], f"KUNCI\nJAWABAN\n\n{jawaban}")
    set_cell_simple(table.rows[12].cells[2], "")
    set_cell_simple(table.rows[13].cells[2], "")
    
    # Row 3 col 3: Question text + options
    soal_text = (
        f"{q['soal']}\n\n"
        f"A. {q['A']}\n\n"
        f"B. {q['B']}\n\n"
        f"C. {q['C']}\n\n"
        f"D. {q['D']}\n\n"
        f"E. {q['E']}"
    )
    set_cell_simple(table.rows[3].cells[3], soal_text)
    
    # Clear rows 4-13 col 3
    for r in range(4, 14):
        try:
            clear_cell_content(table.rows[r].cells[3])
        except Exception:
            pass


def fill_essay_table(table, q):
    """Fill an essay kartu soal table with question data."""
    
    # Row 0 - Header
    fill_header_cell(table.rows[0].cells[0], NAMA_PENULIS, MAPEL, KELAS)
    
    # Row 1-2 col 1-4: Buku Referensi
    for r in [1, 2]:
        for c in [1, 2, 3, 4]:
            try:
                cell = table.rows[r].cells[c]
                clear_cell_content(cell)
                cell.paragraphs[0].add_run(f"Buku Acuan/Referensi:\n{REF}")
            except Exception:
                pass
    
    # Row 2 col 0: Elemen Kompetensi
    set_cell_simple(table.rows[2].cells[0], q['elemen'])
    
    # Row 3 col 0: Materi
    set_cell_simple(table.rows[3].cells[0], q['materi'])
    
    # Row 5 col 0: Capaian Pembelajaran
    set_cell_simple(table.rows[5].cells[0], q['cp'])
    
    # Row 6 col 0
    set_cell_simple(table.rows[6].cells[0], f"Kelas: {KELAS}")
    
    # Row 7 col 0
    set_cell_simple(table.rows[7].cells[0], "Semester: Genap")
    
    # Row 8 col 0
    set_cell_simple(table.rows[8].cells[0], f"Bobot Soal: {q.get('nilai', '10')} poin")
    
    # Row 10 col 0: IPK
    set_cell_simple(table.rows[10].cells[0], q['ipk'])
    
    # Row 12 col 0: Indikator
    set_cell_simple(table.rows[12].cells[0], q['indikator'])
    
    # Row 13 col 0: Tingkat Kesukaran
    tingkat = q.get('tingkat', 'Sedang')
    set_cell_simple(table.rows[13].cells[0],
        f"Tingkat Kesukaran:\n{'✓ Rendah' if tingkat=='Rendah' else '  Rendah'}\n"
        f"{'✓ Sedang' if tingkat=='Sedang' else '  Sedang'}\n"
        f"{'✓ Sulit' if tingkat=='Sulit' else '  Sulit'}")
    
    # Row 5 col 2: No. Soal
    set_cell_simple(table.rows[5].cells[2], str(q['no']))
    
    # Row 8 col 2: Nilai
    set_cell_simple(table.rows[8].cells[2], f"{q.get('nilai', '10')}")
    
    # Row 11-13 col 2: Kunci Jawaban label
    set_cell_simple(table.rows[11].cells[2], "KUNCI\nJAWABAN")
    set_cell_simple(table.rows[12].cells[2], "(lihat\ndi bawah)")
    set_cell_simple(table.rows[13].cells[2], "")
    
    # Row 3 col 3: Narasi + Pertanyaan + Jawaban
    soal_text = (
        f"{q['narasi']}\n\n"
        f"PERTANYAAN:\n{q['pertanyaan']}\n\n"
        f"─────────────────────────────────────────\n"
        f"KUNCI JAWABAN:\n\n{q['jawaban']}"
    )
    set_cell_simple(table.rows[3].cells[3], soal_text)
    
    # Clear rows 4-13 col 3
    for r in range(4, 14):
        try:
            clear_cell_content(table.rows[r].cells[3])
        except Exception:
            pass


# ============================================================
# FILL TABLE 0 (General Header)
# ============================================================

def fill_table0(table):
    cell = table.rows[0].cells[0]
    for para in cell.paragraphs:
        for run in para.runs:
            if '...............................' in run.text:
                run.text = run.text.replace('...............................', NAMA_PENULIS)
            elif '...............................di' in run.text:
                run.text = run.text.replace('...............................di', NAMA_PENULIS)
    # Also try to fix Mata Pelajaran and Kelas
    for para in cell.paragraphs:
        full_text = para.text
        if 'Mata Pelajaran' in full_text and '........' in full_text:
            for run in para.runs:
                if '......' in run.text:
                    run.text = run.text.replace('......', MAPEL if 'Mata' in full_text else KELAS)
        elif 'Kelas' in full_text and '......' in full_text:
            for run in para.runs:
                if '......' in run.text:
                    run.text = run.text.replace('......', KELAS)


# ============================================================
# FILL TABLE 1 (Kisi-kisi)
# ============================================================

def fill_kisi_kisi(table):
    # First section (rows 1-4): PG questions grouped by topic
    pg_groups = [
        ("Pengelolaan Komputer Akuntansi",
         "Pengenalan dan Setup Accurate Online",
         "Peserta didik mampu menjelaskan karakteristik Accurate Online, melakukan setup perusahaan, dan mengelola daftar akun",
         "Mengidentifikasi fitur dan menu Accurate Online untuk pengelolaan data akuntansi",
         "5", "1 – 5", "Pilihan Ganda", "Rendah – Sedang"),
        ("Pengelolaan Komputer Akuntansi",
         "Transaksi Penjualan, Retur Penjualan, dan Penerimaan",
         "Peserta didik mampu mencatat faktur penjualan, retur penjualan, dan penerimaan dari pelanggan",
         "Mencatat siklus transaksi penjualan kredit hingga pelunasan piutang di Accurate Online",
         "5", "6 – 10", "Pilihan Ganda", "Sedang"),
        ("Pengelolaan Komputer Akuntansi",
         "Transaksi Pembelian, Retur Pembelian, dan Pembayaran",
         "Peserta didik mampu mencatat faktur pembelian, retur pembelian, dan pembayaran kepada pemasok",
         "Mencatat siklus transaksi pembelian kredit hingga pelunasan utang di Accurate Online",
         "5", "11 – 15", "Pilihan Ganda", "Sedang"),
        ("Pengelolaan Komputer Akuntansi",
         "Persediaan, Jurnal Umum, dan Laporan Neraca Saldo",
         "Peserta didik mampu melakukan penyesuaian stok, mencatat jurnal umum, dan menghasilkan laporan neraca saldo",
         "Menggunakan fitur Stock Adjustment, General Journal, dan mengakses Trial Balance",
         "5", "16 – 20", "Pilihan Ganda", "Sedang – Sulit"),
    ]
    
    essay_groups = [
        ("Pengelolaan Komputer Akuntansi",
         "Setup Data Awal dan Daftar Akun",
         "Peserta didik mampu menjelaskan prosedur setup awal dan pengelolaan daftar akun di Accurate Online",
         "Mendeskripsikan langkah setup perusahaan dan kelompok akun beserta contohnya",
         "2", "1 – 2", "Uraian/Essay", "Sedang"),
        ("Pengelolaan Komputer Akuntansi",
         "Saldo Awal, Penjualan, dan Retur Penjualan",
         "Peserta didik mampu menjelaskan input saldo awal, prosedur penjualan kredit, dan retur penjualan",
         "Menjelaskan prosedur dan jurnal transaksi penjualan kredit serta retur penjualan",
         "3", "3 – 5", "Uraian/Essay", "Sedang"),
        ("Pengelolaan Komputer Akuntansi",
         "Pembelian, Retur Pembelian, dan Stock Adjustment",
         "Peserta didik mampu menjelaskan pembelian kredit, retur pembelian, dan penyesuaian stok",
         "Menjelaskan prosedur dan jurnal pembelian kredit, retur, dan stock adjustment",
         "3", "6 – 8", "Uraian/Essay", "Sulit"),
        ("Pengelolaan Komputer Akuntansi",
         "Neraca Saldo dan Alur Kerja Accurate Online",
         "Peserta didik mampu menjelaskan neraca saldo dan alur kerja Accurate Online secara komprehensif",
         "Mendeskripsikan fungsi neraca saldo dan workflow Accurate Online dari setup hingga laporan",
         "2", "9 – 10", "Uraian/Essay", "Sulit"),
    ]
    
    # Fill rows 1-4 (PG groups)
    for i, (elemen, materi, cp, tp, jml, nomor, bentuk, tingkat) in enumerate(pg_groups):
        row = table.rows[i + 1]
        data = [str(i+1), elemen, cp, materi, tp, jml, nomor, bentuk, tingkat]
        for c, val in enumerate(data):
            try:
                cell = row.cells[c]
                set_cell_simple(cell, val)
            except Exception:
                pass
    
    # Fill rows 6-9 (Essay groups)
    for i, (elemen, materi, cp, tp, jml, nomor, bentuk, tingkat) in enumerate(essay_groups):
        row = table.rows[i + 6]
        data = [str(i+1), elemen, cp, materi, tp, jml, nomor, bentuk, tingkat]
        for c, val in enumerate(data):
            try:
                cell = row.cells[c]
                set_cell_simple(cell, val)
            except Exception:
                pass
    
    # Fill row 11 (Jumlah total)
    try:
        row = table.rows[11]
        set_cell_simple(row.cells[0], "TOTAL")
        set_cell_simple(row.cells[5], "30")
        set_cell_simple(row.cells[6], "PG: 1-20\nEssay: 1-10")
        set_cell_simple(row.cells[7], "PG & Essay")
        set_cell_simple(row.cells[8], "Variatif")
    except Exception:
        pass


# ============================================================
# MAIN: ADD EXTRA PG TABLES AND FILL ALL
# ============================================================

def main():
    doc = Document('/root/.claude/uploads/19d916f2-492a-4a0c-bfba-2d98a5361dd0/faf676bf-KARTU_SOAL_UasGENAP_2526okeaslikosongan.docx')
    
    print("Adding 7 more PG tables...")
    # Tables 2-14 = 13 PG tables. Need 7 more for questions 14-20.
    # Clone table 14 (last PG table) and insert before table 15 (first essay)
    
    ref_pg_elem = doc.tables[2]._element  # use table 2 as template
    last_pg_elem = doc.tables[14]._element
    parent = last_pg_elem.getparent()
    children = list(parent)
    idx_14 = children.index(last_pg_elem)
    
    # Insert 7 new PG tables after table 14
    for i in range(7):
        new_table = deepcopy(ref_pg_elem)
        parent.insert(idx_14 + 1 + i, new_table)
    
    print(f"Tables now: {len(doc.tables)}")
    
    # Fill Table 0 (general header)
    print("Filling Table 0 (header)...")
    fill_table0(doc.tables[0])
    
    # Fill Table 1 (kisi-kisi)
    print("Filling Table 1 (kisi-kisi)...")
    fill_kisi_kisi(doc.tables[1])
    
    # Fill PG tables (now tables 2-21, for questions 1-20)
    print("Filling PG tables...")
    for i, q in enumerate(pg_data):
        table_idx = i + 2  # tables 2-21
        print(f"  PG Soal {q['no']} -> Table {table_idx}")
        fill_pg_table(doc.tables[table_idx], q)
    
    # Fill Essay tables (tables 22-31, for essay 1-10)
    # After adding 7 tables: old table 15 -> new table 22
    print("Filling Essay tables...")
    for i, q in enumerate(essay_data):
        table_idx = 22 + i  # tables 22-31
        print(f"  Essay Soal {q['no']} -> Table {table_idx}")
        fill_essay_table(doc.tables[table_idx], q)
    
    # Save
    output_path = '/home/user/daw/KARTU_SOAL_UAS_GENAP_KomputerAkuntansi_XI.docx'
    doc.save(output_path)
    print(f"\nSaved to: {output_path}")

main()
