/**
 * ICMS API — MULTI TABLE
 * Google Apps Script + Google Sheets
 *
 * Satu Web App URL ini bisa melayani BANYAK tabel (sheet), dibedakan
 * lewat parameter "sheet" yang dikirim dari frontend. Jika parameter
 * "sheet" tidak dikirim, default-nya adalah "rencana_kelas" (backward
 * compatible dengan index.html yang sudah ada).
 *
 * CARA MENAMBAH TABEL BARU:
 * 1. Tambahkan satu entri baru di object SCHEMAS di bawah ini.
 * 2. Jalankan fungsi setupAllSheets() sekali (Run > setupAllSheets).
 * 3. Di frontend, kirim parameter sheet:"nama_tabel_baru" pada apiGet/apiPost.
 * TIDAK PERLU deploy ulang Web App, dan TIDAK PERLU mengubah kode di bawah.
 *
 * ============================================================
 * CATATAN PERBAIKAN (17 Agustus 2026):
 * 1. save_() sekarang men-trim() nilai id sebelum dibandingkan, agar
 *    pencocokan baris untuk UPDATE tidak gagal karena spasi tersembunyi.
 * 2. importAliases untuk "rekap_program" diperluas (Program/Subprogram)
 *    agar tidak ada lagi kolom kosong akibat header Excel yang variatif.
 * 3. Ditambahkan fungsi utilitas repairMissingIds() untuk memperbaiki
 *    baris-baris LAMA yang kolom id-nya kosong (root cause dari kasus
 *    "edit tidak tersimpan, malah jadi data baru").
 * 4. Skema "summary" DIBANGUN PROGRAMATIS (bukan diketik manual satu-satu)S
 *    dari daftar kolom final yang sudah dipastikan TIDAK ADA NAMA GANDA.
 *    Ini penting karena list_() memetakan baris -> object memakai NAMA
 *    KOLOM sebagai key; kalau ada dua kolom bernama sama persis, data
 *    kolom yang lebih awal akan tertimpa/hilang saat dikirim ke frontend.
 *    Satu nama yang sempat kembar ("TOTAL AKUM" dipakai baik untuk blok
 *    AKUM BLN maupun AKUM OUTBLN) diberi nama unik "TOTAL AKUM OUTBLN"
 *    untuk blok AKUM OUTBLN agar tidak lagi bentrok.
 * ============================================================
 */

const CONFIG = {
  SPREADSHEET_ID: "1d9-y7yUJJ-mPl-NXWdyc-sCgNfph4JvQNNm5zixPm_g",
  API_KEY: "2026-BNU-KARTUKONTROL",
  DEFAULT_SHEET: "rencana_kelas",
  AUTH_SALT: "GANTI-DENGAN-STRING-ACAK-MILIK-SENDIRI-2026",
  SESSION_HOURS: 8,
  SESSIONS_SHEET: "_sessions"
};

/**
 * Helper untuk membangun kolom "summary" secara programatis.
 * Setiap blok bulanan punya 12 kolom (BLN 1..BLN 12 dst) + 1 kolom total.
 * Dibangun lewat loop supaya TIDAK ADA risiko salah ketik/duplikat nama
 * kolom seperti yang terjadi kalau ditulis manual satu per satu.
 */
function buildSummaryMonthBlocks_() {
  return [
    ["BLN", "TOTAL BLN"],
    ["AKUM BLN", "TOTAL AKUM"],
    ["OUTBLN", "TOTAL OUTBLN"],
    ["AKUM OUTBLN", "TOTAL AKUM OUTBLN"], // diganti unik, semula bentrok dgn "TOTAL AKUM"
    ["PROG BLN", "TOTAL PROG"],
    ["SUB PROG BLN", "TOTAL SUB PROG"],
    ["KELAS BLN", "TOTAL KELAS"],
    ["LEARNER BLN", "TOTAL LEARNER"],
    ["RBLN", "RTOTAL BLN"],
    ["RAKUM BLN", "RTOTAL AKUM"],
    ["RPROG BLN", "RTOTAL PROG"],
    ["RSUB PROG BLN", "RTOTAL SUB PROG"],
    ["RKELAS BLN", "RTOTAL KELAS"],
    ["RLEARNER BLN", "RTOTAL LEARNER"]
  ];
}
function buildSummaryBaseFields_() {
  return [
    "Anggaran UNIT", "Target IDR", "Target Proposional IDR", "Realisasi H-1",
    "Pipeline Next Month", "Proyeksi Next Month", "Realisasi Thd Proposional IDR",
    "Realisasi Thd Proposional %", "Realisasi Thd Full Year IDR", "Realisasi Thd Full Year %",
    "Proyeksi Thd Proposional IDR", "Proyeksi Thd Proposional %", "Proyeksi Thd Full Year IDR",
    "Proyeksi Thd Full Year %"
  ];
}
function buildSummaryHeaders_() {
  const out = ["id"].concat(buildSummaryBaseFields_());
  buildSummaryMonthBlocks_().forEach(function (b) {
    const prefix = b[0], totalLabel = b[1];
    for (let i = 1; i <= 12; i++) out.push(prefix + " " + i);
    out.push(totalLabel);
  });
  return out;
}
function buildSummaryNumberFields_() {
  // semua field summary numerik KECUALI "Anggaran UNIT" (nama unit, teks)
  const out = buildSummaryBaseFields_().filter(function (h) { return h !== "Anggaran UNIT"; });
  buildSummaryMonthBlocks_().forEach(function (b) {
    const prefix = b[0], totalLabel = b[1];
    for (let i = 1; i <= 12; i++) out.push(prefix + " " + i);
    out.push(totalLabel);
  });
  return out;
}

/**
 * DEFINISI SKEMA PER TABEL.
 * headers        : urutan kolom persis seperti di sheet (kolom pertama HARUS "id").
 * dateFields     : kolom yang harus diperlakukan sebagai tanggal.
 * numberFields   : kolom yang harus diperlakukan sebagai angka.
 * uniqueField    : (opsional) kolom yang tidak boleh duplikat, mis. "id_class".
 * importAliases  : (opsional) alias nama kolom asli (Excel) -> field key.
 */
const SCHEMAS = {

  rencana_kelas: {
    headers: [
      "id","fileName","id_class","nama_pelatihan","tgl_awal","tgl_akhir","lokasi",
      "npp_lc","pic_lc","pic_kelas","nama","total_input_pegawai","jenis_pelatihan",
      "metode","academy","unit","jumlah_hari","jumlah_peserta_didaftarkan",
      "jumlah_peserta_absen","biaya_pelatihan","status_kelas","tahun","created_at","updated_at"
    ],
    dateFields: ["tgl_awal","tgl_akhir","tahun"],
    numberFields: ["total_input_pegawai","jumlah_hari","jumlah_peserta_didaftarkan","jumlah_peserta_absen","biaya_pelatihan"],
    uniqueField: "id_class",
    importAliases: {
      fileName:["fileName","FILE_NAME","FILE NAME","Kolom 1"],
      id_class:["id_class","IDCLASS","ID CLASS","Kolom 2"],
      nama_pelatihan:["nama_pelatihan","NAMAPELATIHAN","NAMA PELATIHAN","Kolom 3"],
      tgl_awal:["tgl_awal","TGLAWAL","TGL AWAL","Kolom 4"],
      tgl_akhir:["tgl_akhir","TGLAKHIR","TGL AKHIR","Kolom 5"],
      lokasi:["lokasi","LOKASI","Kolom 6"],
      npp_lc:["npp_lc","NPPLC","NPP LC","Kolom 7"],
      pic_lc:["pic_lc","PICLC","PIC LC","Kolom 8"],
      pic_kelas:["pic_kelas","PICKELAS","PIC KELAS","Kolom 9"],
      nama:["nama","NAMA","Kolom 10"],
      total_input_pegawai:["total_input_pegawai","TOTALINPUTPEGAWAI","TOTAL INPUT PEGAWAI","Kolom 11"],
      jenis_pelatihan:["jenis_pelatihan","JENISPELATIHAN","JENIS PELATIHAN","Kolom 12"],
      metode:["metode","METODE","Kolom 13"],
      academy:["academy","ACADEMY","Kolom 14"],
      unit:["unit","UNIT","Kolom 15"],
      jumlah_hari:["jumlah_hari","JUMLAHHARI","JUMLAH HARI","Kolom 16"],
      jumlah_peserta_didaftarkan:["jumlah_peserta_didaftarkan","JUMLAHPESERTADIDAFTARKAN","JUMLAH PESERTA DIDAFTARKAN","Kolom 17"],
      jumlah_peserta_absen:["jumlah_peserta_absen","JUMLAHPESERTAABSEN","JUMLAH PESERTA ABSEN","Kolom 18"],
      biaya_pelatihan:["biaya_pelatihan","BIAYAPELATIHAN","BIAYA PELATIHAN","Kolom 19"],
      status_kelas:["status_kelas","STATUSKELAS","STATUS KELAS","Kolom 20"],
      tahun:["tahun","TAHUN","Kolom 21"]
    }
  },

  learning_plan_realisasi: {
    headers: [
      "id", "uic", "nama_program", "goals", "bln_awal_periode", "bln_akhir_periode",
      "tagging_jenis_anggaran", "nama_account", "natural_account", "latar_belakang_prioritas",
      "klasifikasi_program_1", "klasifikasi_program_2", "klasifikasi_program_3",
      "jenjang_peserta", "divisi_satuan_wilayah_peserta", "nama_posisi_peserta",
      "dewan_komisaris", "direksi", "pejabat_eksekutif", "jabatan_lainnya", "jumlah_peserta",
      "metode_pelaksanaan", "biaya_sertifikasi", "biaya_fasilitator_internal", 
      "biaya_fasilitator_eksternal", "biaya_transportasi", "biaya_akomodasi", 
      "biaya_konsumsi", "biaya_dop", "total_estimasi_biaya", "nama_project_pengadaan",
      "nominal_project_pengadaan", "catatan", "final_nomina",
      "created_at", "updated_at"
    ],
    dateFields: [],
    numberFields: [
      "dewan_komisaris", "direksi", "pejabat_eksekutif", "jabatan_lainnya", "jumlah_peserta",
      "biaya_sertifikasi", "biaya_fasilitator_internal", "biaya_fasilitator_eksternal",
      "biaya_transportasi", "biaya_akomodasi", "biaya_konsumsi", "biaya_dop", 
      "total_estimasi_biaya", "nominal_project_pengadaan", "final_nomina"
    ],
    uniqueField: "",
    importAliases: {
      uic: ["UIC", "Kolom 1"],
      nama_program: ["Nama Program", "Kolom 2"],
      goals: ["Goals", "Kolom 3"],
      bln_awal_periode: ["Bulan Awal Periode Pelaksanaan", "Kolom 4"],
      bln_akhir_periode: ["Bulan Akhir Periode Pelaksanaan", "Kolom 5"],
      tagging_jenis_anggaran: ["Tagging Jenis Anggaran", "Kolom 6"],
      nama_account: ["Nama Account", "Kolom 7"],
      natural_account: ["Natural Account", "Kolom 8"],
      latar_belakang_prioritas: ["Latar Belakang/Prioritas", "Kolom 9"],
      klasifikasi_program_1: ["Klasifikasi Program (1)", "Kolom 10"],
      klasifikasi_program_2: ["Klasifikasi Program (2)", "Kolom 11"],
      klasifikasi_program_3: ["Klasifikasi Program (3)", "Kolom 12"],
      jenjang_peserta: ["Jenjang Peserta", "Kolom 13"],
      divisi_satuan_wilayah_peserta: ["Divisi/Satuan/Wilayah Peserta", "Kolom 14"],
      nama_posisi_peserta: ["Nama Posisi Peserta", "Kolom 15"],
      dewan_komisaris: ["Dewan Komisaris", "Kolom 16"],
      direksi: ["Direksi", "Kolom 17"],
      pejabat_eksekutif: ["Pejabat Eksekutif", "Kolom 18"],
      jabatan_lainnya: ["Jabatan Lainnya", "Kolom 19"],
      jumlah_peserta: ["Jumlah Peserta", "Kolom 20"],
      metode_pelaksanaan: ["Metode Pelaksanaan (Online/Offline/Hybrid)", "Metode Pelaksanaan", "Kolom 21"],
      biaya_sertifikasi: ["Biaya Sertifikasi", "Kolom 22"],
      biaya_fasilitator_internal: ["Biaya Fasilitator Internal", "Kolom 23"],
      biaya_fasilitator_eksternal: ["Biaya Fasilitator Eksternal", "Kolom 24"],
      biaya_transportasi: ["Biaya Transportasi", "Kolom 25"],
      biaya_akomodasi: ["Biaya Akomodasi", "Kolom 26"],
      biaya_konsumsi: ["Biaya Konsumsi", "Kolom 27"],
      biaya_dop: ["Biaya DOP", "Kolom 28"],
      total_estimasi_biaya: ["Total Estimasi Biaya", "Kolom 29"],
      nama_project_pengadaan: ["Nama Project Pengadaan", "Kolom 30"],
      nominal_project_pengadaan: ["Nominal Project Pengadaan", "Kolom 31"],
      catatan: ["Catatan", "Kolom 32"],
      final_nomina: ["FNAL NOMINA", "FINAL NOMINA", "Kolom 33"]
    }
  },

  // ------------------------------------------------------------
  // REKAP PROGRAM — importAliases diperluas agar Program/Subprogram
  // tidak lagi kosong ketika header Excel sumber sedikit berbeda.
  // ------------------------------------------------------------
  rekap_program: {
    headers: [
      "id","uic","program","subprogram","pengelola","tgl_awal","tgl_akhir",
      "metode","jumlah_hari","learner","biaya_pelatihan","created_at","updated_at"
    ],
    dateFields: ["tgl_awal","tgl_akhir"],
    numberFields: ["jumlah_hari","learner","biaya_pelatihan"],
    uniqueField: "",
    importAliases: {
      uic:["UIC","Kolom 1"],
      program:["Program","Nama Program","PROGRAM","Program Pelatihan","Nama Program Pelatihan","Kolom 2"],
      subprogram:["Subprogram","Sub Program","Sub-Program","SUBPROGRAM","Kolom 3"],
      pengelola:["Pengelola","PIC","Nama Pengelola","Kolom 4"],
      tgl_awal:["Tgl Awal","Tanggal Awal","Tgl Mulai","Kolom 5"],
      tgl_akhir:["Tgl Akhir","Tanggal Akhir","Tgl Selesai","Kolom 6"],
      metode:["Metode","METODE","Kolom 7"],
      jumlah_hari:["Jumlah Hari","Jumlah Batch","JUMLAH HARI","JUMLAH BATCH","Kolom 8"],
      learner:["Learner","LEARNER","Jumlah Learner","Kolom 9"],
      biaya_pelatihan:["Biaya Pelatihan","BIAYA PELATIHAN","Kolom 10"]
    }
  },

  users: {
    headers: ["id","username","password_hash","role","active","created_at","updated_at"],
    dateFields: [],
    numberFields: [],
    uniqueField: "username",
    importAliases: {}
  },

  pengeluaran: {
    headers: [
      "id","fileName","no_pp","create_pp","no_voucer","tgl_voucer","print_pp","nama_pp",
      "program","unit","miscode","nama_program","course_title","nama_rekening_debet","nilai",
      "cost_code","nama_cost_code","status_pp","id_class","id_sentralisasi","status_lokasi","uic",
      "created_at","updated_at"
    ],
    dateFields: ["create_pp","tgl_voucer"],
    numberFields: ["nilai"],
    uniqueField: "",
    importAliases: {
      fileName:["fileName","FILE_NAME","FILE NAME", "Kolom 1"],
      no_pp:["no_pp","No PP","NO_PP", "Kolom 2"],
      create_pp:["create_pp","Create PP","CREATE_PP", "Kolom 3"],
      no_voucer:["no_voucer","No Voucer","NO_VOUCER", "Kolom 4"],
      tgl_voucer:["tgl_voucer","Tgl Voucer","TGL_VOUCER", "Kolom 5"],
      print_pp:["print_pp","Print_pp","Print PP","PRINT_PP", "Kolom 6"],
      nama_pp:["nama_pp","Nama PP","NAMA_PP", "Kolom 7"],
      program:["program","Program", "Kolom 8"],
      unit:["unit","Unit", "Kolom 9"],
      miscode:["miscode","Miscode","MISCODE", "Kolom 10"],
      nama_program:["nama_program","Nama Program", "Kolom 11"],
      course_title:["course_title","Course Title", "Kolom 12"],
      nama_rekening_debet:["nama_rekening_debet","Nama Rekening Debet", "Kolom 13"],
      nilai:["nilai","Nilai", "Kolom 14"],
      cost_code:["cost_code","Cost Code", "Kolom 15"],
      nama_cost_code:["nama_cost_code","Nama Cost Code", "Kolom 16"],
      status_pp:["status_pp","Status PP", "Kolom 17"],
      id_class:["id_class","ID Class","ID_CLASS", "Kolom 18"],
      id_sentralisasi:["id_sentralisasi","ID Sentralisasi", "Kolom 19"],
      status_lokasi:["status_lokasi","Status Lokasi", "Kolom 20"],
      uic:["uic","UIC", "Kolom 21"]
    }
  },

  keu_smarter: {
    headers: [
      "id","fileName","no_pp","create_pp","no_voucer","tgl_voucer","print_pp","nama_pp",
      "program","unit","miscode","nama_program","course_title","debet","nama_rekening_debet",
      "kredit","nama_rekening_kredit","nilai","keterangan","cost_code","nama_cost_code",
      "status_pp","id_class","id_sentralisasi","status_lokasi","pengajuan_biaya",
      "biaya_pelatihan","kelompok","alokasi","saldo","tgl_report","created_at","updated_at"
    ],
    dateFields: ["create_pp","tgl_voucer","tgl_report"],
    numberFields: ["debet","kredit","nilai","pengajuan_biaya","biaya_pelatihan","alokasi","saldo"],
    uniqueField: "",
    importAliases: {
      fileName:["fileName","FILE_NAME","FILE NAME"],
      no_pp:["no_pp","No PP","NO_PP"],
      create_pp:["create_pp","Create PP","CREATE_PP"],
      no_voucer:["no_voucer","No Voucer","NO_VOUCER"],
      tgl_voucer:["tgl_voucer","Tgl Voucer","TGL_VOUCER"],
      print_pp:["print_pp","Print_pp","Print PP","PRINT_PP"],
      nama_pp:["nama_pp","Nama PP","NAMA_PP"],
      program:["program","Program"],
      unit:["unit","Unit"],
      miscode:["miscode","Miscode","MISCODE"],
      nama_program:["nama_program","Nama Program"],
      course_title:["course_title","Course Title"],
      debet:["debet","Debet"],
      nama_rekening_debet:["nama_rekening_debet","Nama Rekening Debet"],
      kredit:["kredit","Kredit"],
      nama_rekening_kredit:["nama_rekening_kredit","Nama Rekening Kredit"],
      nilai:["nilai","Nilai"],
      keterangan:["keterangan","Keterangan"],
      cost_code:["cost_code","Cost Code"],
      nama_cost_code:["nama_cost_code","Nama Cost Code"],
      status_pp:["status_pp","Status PP"],
      id_class:["id_class","ID Class","ID_CLASS"],
      id_sentralisasi:["id_sentralisasi","ID Sentralisasi"],
      status_lokasi:["status_lokasi","Status Lokasi"],
      pengajuan_biaya:["pengajuan_biaya","Pengajuan Biaya"],
      biaya_pelatihan:["biaya_pelatihan","Biaya Pelatihan"],
      kelompok:["kelompok","Kelompok"],
      alokasi:["alokasi","Alokasi"],
      saldo:["saldo","Saldo"],
      tgl_report:["tgl_report","Tgl Report"]
    }
  },

  keu_portal: {
    headers: [
      "id","fileName","src_id","no_pp","no_voucer","pic","nama_pp","program","tgl_voucer",
      "miscode","created_by","create_date","update_date","update_by","keterangan","debet",
      "nama_rekening_debet","kredit","nama_rekening_kredit","print_flag","terima_flag",
      "reject_flag","tanggal_buku","kn_flag","nilai","status_pp","status_lokasi","tgl_report",
      "created_at","updated_at"
    ],
    dateFields: ["tgl_voucer","create_date","update_date","tanggal_buku","tgl_report"],
    numberFields: ["debet","kredit","nilai"],
    uniqueField: "",
    importAliases: {
      fileName:["fileName","FILE_NAME","FILE NAME"],
      src_id:["src_id","id","ID"],
      no_pp:["no_pp","No PP","NO_PP"],
      no_voucer:["no_voucer","No Voucer","NO_VOUCER"],
      pic:["pic","PIC"],
      nama_pp:["nama_pp","Nama PP","NAMA_PP"],
      program:["program","Program"],
      tgl_voucer:["tgl_voucer","Tgl Voucer","TGL_VOUCER"],
      miscode:["miscode","Miscode","MISCODE"],
      created_by:["created_by","Created By"],
      create_date:["create_date","Create Date"],
      update_date:["update_date","Update Date"],
      update_by:["update_by","Update By"],
      keterangan:["keterangan","Keterangan"],
      debet:["debet","Debet"],
      nama_rekening_debet:["nama_rekening_debet","Nama Rekening Debet"],
      kredit:["kredit","Kredit"],
      nama_rekening_kredit:["nama_rekening_kredit","Nama Rekening Kredit"],
      print_flag:["print_flag","Print Flag"],
      terima_flag:["terima_flag","Terima Flag"],
      reject_flag:["reject_flag","Reject Flag"],
      tanggal_buku:["tanggal_buku","Tanggal Buku"],
      kn_flag:["kn_flag","KN Flag"],
      nilai:["nilai","Nilai"],
      status_pp:["status_pp","Status PP"],
      status_lokasi:["status_lokasi","Status Lokasi"],
      tgl_report:["tgl_report","Tgl Report"]
    }
  },

  // ------------------------------------------------------------
  // SUMMARY — headers & numberFields dibangun PROGRAMATIS lewat
  // buildSummaryHeaders_() / buildSummaryNumberFields_() di atas,
  // supaya tidak ada nama kolom yang bentrok (lihat catatan di
  // bagian atas file ini). Kolom pertama tetap "id".
  // ------------------------------------------------------------
  summary: {
    headers: buildSummaryHeaders_(),
    dateFields: [],
    numberFields: buildSummaryNumberFields_(),
    uniqueField: "Anggaran UNIT",
    importAliases: {}
  },

  schedules: {
    headers: ["id", "schedule_date", "title", "description", "color", "status", "created_at", "updated_at"],
    dateFields: ["schedule_date"],
    numberFields: [],
    uniqueField: "",
    importAliases: {}
  }

};

function setupAllSheets(){
  const results=[];
  Object.keys(SCHEMAS).forEach(name=>results.push(setupSheet_(name)));
  return results.join(" | ");
}

function setupSheet_(sheetName){
  const schema=getSchema_(sheetName);
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh=ss.getSheetByName(sheetName);
  if(!sh) sh=ss.insertSheet(sheetName);
  sh.clear();
  sh.getRange(1,1,1,schema.headers.length).setValues([schema.headers]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,schema.headers.length).setFontWeight("bold");
  return "Sheet siap: "+sheetName;
}

function getSchema_(sheetName){
  const schema=SCHEMAS[sheetName];
  if(!schema) throw new Error("Tabel/sheet tidak dikenal: "+sheetName);
  return schema;
}

function doGet(e){
  try{
    auth_(e.parameter.key);
    const action = e.parameter.action || "list";
    const sheetName = e.parameter.sheet || CONFIG.DEFAULT_SHEET;
    if(action==="checkSession"){
      const u=validateSession_(e.parameter.token);
      return json_({ok:true,username:u.username,role:u.role});
    }
    const session=requireSession_(e.parameter.token);
    if(sheetName==="users" && session.role!=="admin") throw new Error("Hanya admin yang dapat melihat data user.");
    if(action==="list") return json_({ok:true,data:list_(sheetName)});
    if(action==="setup") return json_({ok:true,msg:setupSheet_(sheetName)});
    if(action==="setupAll") return json_({ok:true,msg:setupAllSheets()});
    if(action==="tables") return json_({ok:true,data:Object.keys(SCHEMAS)});
    return json_({ok:false,msg:"Action GET tidak dikenal."});
  }catch(err){return json_({ok:false,msg:String(err.message||err)})}
}

function doPost(e){
  try{
    const p = e.parameter || {};
    auth_(p.key);
    const action = p.action || "";
    const sheetName = p.sheet || CONFIG.DEFAULT_SHEET;
    if(action==="login") return json_(login_(p));
    if(action==="logout"){ destroySession_(p.token); return json_({ok:true,msg:"Logout berhasil."}); }
    const session=requireSession_(p.token);
    if(sheetName==="users" && session.role!=="admin") throw new Error("Hanya admin yang dapat mengelola user.");
    if(action==="save") return json_(save_(sheetName,p));
    if(action==="delete") return json_(delete_(sheetName,p.id));
    if(action==="deleteAll") return json_(deleteAll_(sheetName));
    if(action==="import") return json_(import_(sheetName,p.rows));
    return json_({ok:false,msg:"Action POST tidak dikenal."});
  }catch(err){return json_({ok:false,msg:String(err.message||err)})}
}

function auth_(key){
  if(CONFIG.API_KEY && key !== CONFIG.API_KEY) throw new Error("API key tidak valid.");
}

// ============================================================
// AUTH: hashing, session, login/logout
// ============================================================

function hashPassword_(pw){
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw)+CONFIG.AUTH_SALT);
  return bytes.map(b=>((b<0?b+256:b).toString(16)).padStart(2,"0")).join("");
}

function sessionsSheet_(){
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh=ss.getSheetByName(CONFIG.SESSIONS_SHEET);
  if(!sh){
    sh=ss.insertSheet(CONFIG.SESSIONS_SHEET);
    sh.getRange(1,1,1,5).setValues([["token","username","role","created_at","expires_at"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function createSession_(username,role){
  const sh=sessionsSheet_();
  const token=Utilities.getUuid();
  const now=new Date();
  const expires=new Date(now.getTime()+CONFIG.SESSION_HOURS*3600*1000);
  sh.appendRow([token,username,role,now,expires]);
  return {token,expires};
}

function validateSession_(token){
  if(!token) throw new Error("Sesi tidak ditemukan. Silakan login.");
  const sh=sessionsSheet_();
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(token)){
      const expires=new Date(rows[i][4]);
      if(isNaN(expires.getTime())||expires.getTime()<Date.now()){
        sh.deleteRow(i+1);
        throw new Error("Sesi kadaluarsa. Silakan login ulang.");
      }
      return {username:rows[i][1], role:rows[i][2]};
    }
  }
  throw new Error("Sesi tidak valid. Silakan login ulang.");
}

function requireSession_(token){
  return validateSession_(token);
}

function destroySession_(token){
  if(!token)return;
  const sh=sessionsSheet_();
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(token)){ sh.deleteRow(i+1); return; }
  }
}

function login_(p){
  const username=String(p.username||"").trim();
  const password=String(p.password||"");
  if(!username||!password) throw new Error("Username dan password wajib diisi.");
  const sh=sheet_("users");
  const rows=sh.getDataRange().getValues();
  const headers=rows[0];
  const idxUser=headers.indexOf("username");
  const idxHash=headers.indexOf("password_hash");
  const idxRole=headers.indexOf("role");
  const idxActive=headers.indexOf("active");
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][idxUser]).trim().toLowerCase()===username.toLowerCase()){
      const activeVal=String(rows[i][idxActive]).trim().toLowerCase();
      if(activeVal==="false"||activeVal==="0"||activeVal==="tidak"||activeVal==="nonaktif"){
        throw new Error("Akun ini dinonaktifkan. Hubungi admin.");
      }
      if(String(rows[i][idxHash])!==hashPassword_(password)){
        throw new Error("Username atau password salah.");
      }
      const role=rows[i][idxRole]||"user";
      const session=createSession_(username,role);
      return {ok:true,token:session.token,username,role,expires:session.expires};
    }
  }
  throw new Error("Username atau password salah.");
}

/**
 * ============================================================
 * CARA PAKAI (jalankan SEKALI SAJA, saat setup awal sistem):
 * 1. Ganti nilai username & password di bawah ini sesuai keinginan.
 * 2. Di editor Apps Script, pilih fungsi "buatAdminPertama" dari
 *    dropdown fungsi (di sebelah tombol Run/▶), lalu klik Run.
 * 3. Buka tab "Execution log" (Ctrl+Enter) untuk melihat hasilnya.
 * 4. Setelah berhasil, langsung login di aplikasi dengan akun ini,
 *    lalu SEGERA ganti passwordnya lewat halaman "Kelola User".
 * 5. Setelah admin pertama berhasil dibuat, kosongkan/ganti lagi
 *    nilai password di bawah ini agar tidak tertinggal di kode.
 * ============================================================
 */
function buatAdminPertama(){
  const username="user";
  const password="budibaik";

  if(!username || !password || password.length<8){
    Logger.log("Gagal: username wajib diisi dan password minimal 8 karakter.");
    return "Gagal: username wajib diisi dan password minimal 8 karakter.";
  }

  const sh=sheet_("users"); // otomatis membuat sheet 'users' + header jika belum ada
  const rows=sh.getDataRange().getValues();
  const headers=rows[0];
  const idxUser=headers.indexOf("username");
  const idxRole=headers.indexOf("role");

  // Cegah duplikat username yang sama.
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][idxUser]).trim().toLowerCase()===username.toLowerCase()){
      const msg="User '"+username+"' sudah ada, tidak dibuat ulang.";
      Logger.log(msg);
      return msg;
    }
  }

  // Info saja (tidak menghentikan proses): beri tahu jika sudah ada admin lain.
  const adminExists = rows.slice(1).some(r=>String(r[idxRole]).trim().toLowerCase()==="admin");
  if(adminExists){
    Logger.log("Catatan: sudah ada admin lain di sheet 'users'. Tetap membuat admin baru: "+username);
  }

  const now=new Date();
  sh.appendRow([Utilities.getUuid(), username, hashPassword_(password), "admin", true, now, now]);

  const result="✅ Admin pertama berhasil dibuat.\nUsername: "+username+"\nSegera login lalu ganti password lewat halaman Kelola User.";
  Logger.log(result);
  return result;
}

/**
 * ============================================================
 * UTILITAS PERBAIKAN DATA — repairMissingIds
 * ============================================================
 * Fungsi ini memperbaiki baris-baris LAMA yang kolom "id"-nya kosong
 * (biasanya akibat proses import/edit manual sebelumnya di luar sistem).
 * Baris dengan id kosong menyebabkan tombol Edit gagal mencocokkan
 * baris yang tepat, sehingga sistem membuat baris BARU alih-alih
 * mengupdate baris yang sedang diedit.
 *
 * CARA PAKAI:
 * 1. Di editor Apps Script, pilih fungsi "repairMissingIds" dari
 *    dropdown fungsi, lalu klik Run.
 * 2. Secara default akan memperbaiki sheet "rekap_program". Untuk
 *    memperbaiki sheet lain, ubah nama di dalam pemanggilan di baris
 *    terakhir fungsi ini, atau jalankan repairMissingIds("nama_sheet")
 *    langsung dari Execution log / dari fungsi lain.
 * 3. Lihat hasilnya di Execution log (Ctrl+Enter setelah Run).
 * 4. SELALU buat salinan/backup Spreadsheet sebelum menjalankan ini,
 *    sebagai praktik pengamanan data standar.
 * ============================================================
 */
function repairMissingIds(sheetNameInput){
  const sheetName = sheetNameInput || "rekap_program";
  getSchema_(sheetName); // validasi nama tabel dikenal (lempar error jika tidak dikenal)

  // PENTING: pakai getSheetByName langsung (BUKAN sheet_()), agar fungsi ini
  // TIDAK secara tidak sengaja membuat sheet kosong baru untuk tabel yang
  // memang belum pernah dipakai (relevan saat dipanggil dari repairMissingIdsAll
  // yang melewati semua nama tabel di SCHEMAS).
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if(!sh){
    const msg = "Sheet '"+sheetName+"' belum ada di Spreadsheet, dilewati (tidak dibuat otomatis oleh fungsi ini).";
    Logger.log(msg);
    return msg;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if(lastRow<2){
    const msg = "Tidak ada data di sheet: "+sheetName;
    Logger.log(msg);
    return msg;
  }

  // Baca seluruh data (tanpa header) dalam SATU kali panggilan.
  const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  let fixed=0;
  const newIdColumn = data.map(row=>{
    const idVal = String(row[0]||"").trim();
    const hasOtherData = row.slice(1).some(v=>String(v||"").trim()!=="");
    if(!idVal && hasOtherData){
      fixed++;
      return [Utilities.getUuid()];
    }
    return [row[0]];
  });

  // Tulis balik kolom id dalam SATU kali panggilan (jauh lebih cepat &
  // lebih aman terhadap timeout dibanding menulis sel satu per satu,
  // terutama untuk sheet dengan ribuan baris).
  if(fixed>0){
    sh.getRange(2,1,newIdColumn.length,1).setValues(newIdColumn);
  }

  const msg = "Selesai. "+fixed+" dari "+data.length+" baris di sheet '"+sheetName+"' diberi ID baru.";
  Logger.log(msg);
  return msg;
}

/**
 * ============================================================
 * repairMissingIdsAll — menjalankan repairMissingIds() ke SEMUA
 * sheet yang terdaftar di SCHEMAS, satu per satu, dan mencatat
 * ringkasan hasilnya. Sheet internal "_sessions" tidak termasuk
 * karena tidak terdaftar di SCHEMAS.
 *
 * CATATAN: fungsi ini HANYA mengisi kolom id yang KOSONG pada
 * baris yang sudah punya data lain. Baris dengan id yang sudah
 * terisi TIDAK akan diubah/ditimpa.
 *
 * Untuk sheet dengan jumlah baris sangat besar (puluhan ribu),
 * pertimbangkan menjalankan repairMissingIds() per-sheet secara
 * terpisah agar tidak berisiko timeout (batas eksekusi GAS 6 menit).
 * ============================================================
 */
function repairMissingIdsAll(){
  const results=[];
  Object.keys(SCHEMAS).forEach(name=>{
    try{
      results.push(repairMissingIds(name));
    }catch(err){
      results.push("Gagal pada sheet '"+name+"': "+err.message);
    }
  });
  const summary = results.join("\n");
  Logger.log(summary);
  return summary;
}

function sheet_(sheetName){
  getSchema_(sheetName); // validasi tabel dikenal
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh=ss.getSheetByName(sheetName);
  if(!sh){setupSheet_(sheetName);sh=ss.getSheetByName(sheetName);}
  return sh;
}

function list_(sheetName){
  const sh=sheet_(sheetName);
  const values=sh.getDataRange().getValues();
  if(values.length<2)return [];
  const headers=values[0];
  const rows=values.slice(1).filter(r=>r.some(v=>v!=="")).map(r=>{
    const o={};
    headers.forEach((h,i)=>o[h]=normalize_(r[i]));
    return o;
  });
  if(sheetName==="users") rows.forEach(o=>{ delete o.password_hash; });
  return rows;
}

function save_(sheetName,p){
  const schema=getSchema_(sheetName);
  const sh=sheet_(sheetName);
  const now=new Date();
  // PERBAIKAN: trim() id yang datang dari frontend dan dari sheet,
  // agar pencocokan baris untuk UPDATE tidak gagal karena spasi
  // tersembunyi atau perbedaan format string.
  const id=String(p.id||"").trim() || Utilities.getUuid();
  const rows=sh.getDataRange().getValues();
  const headers=rows[0];
  let rowIndex=-1;
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0]).trim()===String(id).trim()){rowIndex=i+1;break;}
  }

  // Validasi kolom unik (jika didefinisikan di schema)
  if(schema.uniqueField){
    const uniqueColIdx=headers.indexOf(schema.uniqueField);
    const uniqueVal=String(p[schema.uniqueField]||"").trim();
    if(uniqueVal && uniqueColIdx>-1){
      for(let i=1;i<rows.length;i++){
        if(rowIndex===i+1)continue;
        if(String(rows[i][uniqueColIdx]).trim()===uniqueVal){
          throw new Error(schema.uniqueField+" sudah digunakan: "+uniqueVal);
        }
      }
    }
  }

  const createdAtIdx=headers.indexOf("created_at");
  const old = rowIndex>0 ? rows[rowIndex-1] : [];
  const createdAt = rowIndex>0 && createdAtIdx>-1 ? old[createdAtIdx] : now;

  // Untuk tabel users: hash password sebelum disimpan, jangan pernah simpan teks asli.
  if(sheetName==="users"){
    if(p.password){
      p.password_hash = hashPassword_(p.password);
    }else if(rowIndex>0){
      const hashIdx=headers.indexOf("password_hash");
      p.password_hash = hashIdx>-1 ? old[hashIdx] : "";
    }else{
      throw new Error("Password wajib diisi untuk user baru.");
    }
    if(p.active===undefined || p.active==="") p.active = true;
  }

  const obj={id:id};
  headers.forEach(h=>{
    if(h==="id")return;
    if(h==="created_at"){obj[h]=createdAt;return;}
    if(h==="updated_at"){obj[h]=now;return;}
    if(schema.dateFields.includes(h)){obj[h]=parseDate_(p[h]);return;}
    if(schema.numberFields.includes(h)){obj[h]=num_(p[h]);return;}
    obj[h]=p[h]||"";
  });

  const row=headers.map(h=>obj[h]??"");
  if(rowIndex>0)sh.getRange(rowIndex,1,1,headers.length).setValues([row]);
  else sh.appendRow(row);
  return {ok:true,msg:rowIndex>0?"Data berhasil diperbarui.":"Data berhasil ditambahkan.",id:id};
}

function delete_(sheetName,id){
  if(!id)throw new Error("ID tidak diberikan.");
  const sh=sheet_(sheetName), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0]).trim()===String(id).trim()){sh.deleteRow(i+1);return {ok:true,msg:"Data berhasil dihapus."};}
  }
  throw new Error("Data tidak ditemukan.");
}

function deleteAll_(sheetName){
  const sh=sheet_(sheetName);
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  return {ok:true,msg:"Semua data berhasil dihapus."};
}

function import_(sheetName,raw){
  const schema=getSchema_(sheetName);
  if(!raw)throw new Error("Data import kosong.");
  const rows=JSON.parse(raw);
  if(!Array.isArray(rows)||!rows.length)throw new Error("Data import tidak valid.");
  let ok=0,fail=[];
  rows.forEach((r,i)=>{
    try{
      const p=mapImport_(schema,r);
      save_(sheetName,p);ok++;
    }catch(err){fail.push("Baris "+(i+2)+": "+err.message)}
  });
  return {ok:true,msg:ok+" data berhasil diimpor."+ (fail.length?" "+fail.length+" gagal.":""),inserted_count:ok,failed_count:fail.length,errors:fail};
}

function mapImport_(schema,r){
  const aliases=schema.importAliases||{};
  const get=(names)=>{
    for(const n of names){
      const key=Object.keys(r).find(k=>normalizeHeader_(k)===normalizeHeader_(n));
      if(key!==undefined)return r[key];
    }
    return "";
  };
  const p={id:""};
  schema.headers.forEach(h=>{
    if(h==="id"||h==="created_at"||h==="updated_at")return;
    const names=aliases[h] || [h];
    p[h]=get(names);
  });
  return p;
}

function normalizeHeader_(s){return String(s||"").toUpperCase().replace(/[^A-Z0-9]/g,"");}
function num_(v){if(v===""||v===null||v===undefined)return 0;const n=Number(String(v).replace(/[^0-9.-]/g,""));return isNaN(n)?0:n}
function parseDate_(v){
  if(!v)return "";
  if(Object.prototype.toString.call(v)==="[object Date]")return v;
  const s=String(v).trim();
  const d=new Date(s);
  if(!isNaN(d.getTime()))return d;
  const m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if(m)return new Date(Number(m[3]),Number(m[2])-1,Number(m[1]));
  return s;
}
function normalize_(v){
  if(v===null||v===undefined)return "";
  if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm:ss");
  return v;
}
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}