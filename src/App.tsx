import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import SignatureCanvas from "react-signature-canvas";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ChartData,
  ChartOptions,
  ChartEvent,
  LegendItem,
} from "chart.js";
import { Bar } from "react-chartjs-2";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

const endpoint =
  "https://script.google.com/macros/s/AKfycbxB1WO7DXWe6--9QXf-IlZoZlyZAvCtnrrfp5vGhX_W-nbOkGne1JuYnz7vl_v4pD-U/exec";
const SHEET_SEMESTER1 = "RekapSemester1";
const SHEET_SEMESTER2 = "RekapSemester2";

interface Student {
  id: string;
  name: string | null | undefined;
  nisn: string | null | undefined;
  kelas: string | null | undefined;
}

interface SchoolData {
  namaKepsek: string;
  nipKepsek: string;
  ttdKepsek: string;
  namaGuru: string;
  nipGuru: string;
  ttdGuru: string;
}

type AttendanceStatus = "Hadir" | "Izin" | "Sakit" | "Alpha";

interface AttendanceRecord {
  [date: string]: {
    [studentId: string]: AttendanceStatus;
  };
}

interface MonthlyRecap {
  nama: string;
  kelas: string;
  hadir: number;
  alpa: number;
  izin: number;
  sakit: number;
  persenHadir: number;
}

interface GraphData {
  [month: string]: {
    Hadir: number;
    Alpha: number;
    Izin: number;
    Sakit: number;
  };
}

interface StatusSummary {
  Hadir: number;
  Izin: number;
  Sakit: number;
  Alpha: number;
}

interface StatusVisibility {
  Hadir: boolean;
  Alpha: boolean;
  Izin: boolean;
  Sakit: boolean;
}

interface AttendanceHistory {
  tanggal: string;
  nama: string;
  kelas: string;
  nisn: string;
  status: AttendanceStatus;
  duplikat: string;
}

interface SemesterRecap {
  nama: string;
  kelas: string;
  hadir: number;
  alpa: number;
  izin: number;
  sakit: number;
  persenHadir: number;
}

const formatDateDDMMYYYY = (isoDate: string): string => {
  const [year, month, day] = isoDate.split("-");
  return `${day}/${month}/${year}`;
};

const SchoolDataTab: React.FC<{
  onRefresh: () => void;
}> = ({ onRefresh }) => {
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [namaKepsek, setNamaKepsek] = useState("");
  const [nipKepsek, setNipKepsek] = useState("");
  const [namaGuru, setNamaGuru] = useState("");
  const [nipGuru, setNipGuru] = useState("");
  const [ttdKepsek, setTtdKepsek] = useState("");
  const [ttdGuru, setTtdGuru] = useState("");
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false); // Add isSaving state
  const [isKepsekSigning, setIsKepsekSigning] = useState(false);
  const [isGuruSigning, setIsGuruSigning] = useState(false);
  const kepsekSigCanvas = useRef<SignatureCanvas>(null);
  const guruSigCanvas = useRef<SignatureCanvas>(null);

  useEffect(() => {
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          const record = data.data[0];
          setSchoolData(record);
          setNamaKepsek(record.namaKepsek);
          setNipKepsek(record.nipKepsek);
          setTtdKepsek(record.ttdKepsek);
          setNamaGuru(record.namaGuru);
          setNipGuru(record.nipGuru);
          setTtdGuru(record.ttdGuru);
        } else {
          setSchoolData(null);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
        setLoading(false);
      });
  }, []);

  const handleSave = () => {
    if (!namaKepsek || !nipKepsek || !namaGuru || !nipGuru) {
      alert("⚠️ Semua field wajib diisi kecuali tanda tangan!");
      return;
    }

    setIsSaving(true); // Set saving state to true

    const data: SchoolData = {
      namaKepsek,
      nipKepsek,
      ttdKepsek: ttdKepsek || "",
      namaGuru,
      nipGuru,
      ttdGuru: ttdGuru || "",
    };

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "schoolData",
        ...data,
      }),
    })
      .then(() => {
        alert("✅ Data sekolah berhasil diperbarui!");
        onRefresh();
        setIsSaving(false); // Reset saving state on success
      })
      .catch(() => {
        alert("❌ Gagal memperbarui data sekolah.");
        setIsSaving(false); // Reset saving state on error
      });
  };

  const handleClearKepsekSignature = () => {
    kepsekSigCanvas.current?.clear();
  };

  const handleClearGuruSignature = () => {
    guruSigCanvas.current?.clear();
  };

  const handleSaveKepsekSignature = () => {
    const signature = kepsekSigCanvas.current?.toDataURL("image/png");
    if (signature && !kepsekSigCanvas.current?.isEmpty()) {
      setTtdKepsek(signature);
      setIsKepsekSigning(false);
    } else {
      alert("⚠️ Tanda tangan kepala sekolah kosong!");
    }
  };

  const handleSaveGuruSignature = () => {
    const signature = guruSigCanvas.current?.toDataURL("image/png");
    if (signature && !guruSigCanvas.current?.isEmpty()) {
      setTtdGuru(signature);
      setIsGuruSigning(false);
    } else {
      alert("⚠️ Tanda tangan guru kosong!");
    }
  };

  const handleStartKepsekSigning = () => {
    setIsKepsekSigning(true);
    kepsekSigCanvas.current?.clear();
  };

  const handleStartGuruSigning = () => {
    setIsGuruSigning(true);
    guruSigCanvas.current?.clear();
  };

  if (loading) {
    return (
      <div className="text-center py-8">
        <p className="text-gray-500">Memuat data sekolah...</p>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          🏫 Data Sekolah
        </h2>
        <div className="grid grid-cols-1 gap-4 mb-6">
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">
              Kepala Sekolah
            </h3>
            <input
              type="text"
              placeholder="Nama Kepala Sekolah"
              value={namaKepsek}
              onChange={(e) => setNamaKepsek(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving} // Disable input during saving
            />
            <input
              type="text"
              placeholder="NIP Kepala Sekolah"
              value={nipKepsek}
              onChange={(e) => setNipKepsek(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving} // Disable input during saving
            />
            <div className="mb-2">
              <p className="text-sm text-gray-500 mb-1">
                Tanda Tangan Kepala Sekolah
              </p>
              <div className="relative">
                <SignatureCanvas
                  ref={kepsekSigCanvas}
                  penColor="black"
                  canvasProps={{
                    className: `border border-gray-300 rounded-lg ${
                      !isKepsekSigning || isSaving
                        ? "opacity-50 pointer-events-none"
                        : ""
                    }`,
                    style: { width: "100%", height: "300px" },
                  }}
                  clearOnResize={false}
                />
                {!isKepsekSigning && (
                  <div className="absolute inset-0 flex items-center justify-center bg-gray-200 bg-opacity-50">
                    <span className="text-gray-500">
                      Klik "Mulai Tanda Tangan" untuk mengaktifkan
                    </span>
                  </div>
                )}
              </div>
              <div className="flex gap-2 mt-2">
                {!isKepsekSigning && (
                  <button
                    onClick={handleStartKepsekSigning}
                    className={`px-4 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    ✍️ Mulai Tanda Tangan
                  </button>
                )}
                {isKepsekSigning && (
                  <button
                    onClick={handleSaveKepsekSignature}
                    className={`px-4 py-1 bg-blue-500 hover:bg-blue-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    💾 Simpan Tanda Tangan
                  </button>
                )}
                <button
                  onClick={handleClearKepsekSignature}
                  className={`px-4 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-sm ${
                    !isKepsekSigning || isSaving
                      ? "opacity-50 cursor-not-allowed"
                      : ""
                  }`}
                  disabled={!isKepsekSigning || isSaving} // Disable during saving
                >
                  🗑️ Hapus TTD
                </button>
              </div>
            </div>
            {ttdKepsek && (
              <img
                src={ttdKepsek}
                alt="Tanda Tangan Kepala Sekolah"
                className="mt-2 max-w-full h-20 border border-gray-200 rounded-lg"
              />
            )}
          </div>
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">Guru</h3>
            <input
              type="text"
              placeholder="Nama Guru"
              value={namaGuru}
              onChange={(e) => setNamaGuru(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving} // Disable input during saving
            />
            <input
              type="text"
              placeholder="NIP Guru"
              value={nipGuru}
              onChange={(e) => setNipGuru(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving} // Disable input during saving
            />
            <div className="mb-2">
              <p className="text-sm text-gray-500 mb-1">Tanda Tangan Guru</p>
              <div className="relative">
                <SignatureCanvas
                  ref={guruSigCanvas}
                  penColor="black"
                  canvasProps={{
                    className: `border border-gray-300 rounded-lg ${
                      !isGuruSigning || isSaving
                        ? "opacity-50 pointer-events-none"
                        : ""
                    }`,
                    style: { width: "100%", height: "300px" },
                  }}
                  clearOnResize={false}
                />
                {!isGuruSigning && (
                  <div className="absolute inset-0 flex items-center justify-center bg-gray-200 bg-opacity-50">
                    <span className="text-gray-500">
                      Klik "Mulai Tanda Tangan" untuk mengaktifkan
                    </span>
                  </div>
                )}
              </div>
              <div className="flex gap-2 mt-2">
                {!isGuruSigning && (
                  <button
                    onClick={handleStartGuruSigning}
                    className={`px-4 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    ✍️ Mulai Tanda Tangan
                  </button>
                )}
                {isGuruSigning && (
                  <button
                    onClick={handleSaveGuruSignature}
                    className={`px-4 py-1 bg-blue-500 hover:bg-blue-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    💾 Simpan Tanda Tangan
                  </button>
                )}
                <button
                  onClick={handleClearGuruSignature}
                  className={`px-4 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-sm ${
                    !isGuruSigning || isSaving
                      ? "opacity-50 cursor-not-allowed"
                      : ""
                  }`}
                  disabled={!isGuruSigning || isSaving} // Disable during saving
                >
                  🗑️ Hapus TTD
                </button>
              </div>
            </div>
            {ttdGuru && (
              <img
                src={ttdGuru}
                alt="Tanda Tangan Guru"
                className="mt-2 max-w-full h-20 border border-gray-200 rounded-lg"
              />
            )}
          </div>
        </div>
        <div className="text-center">
          <button
            onClick={handleSave}
            disabled={isSaving} // Disable button during saving
            className={`px-6 py-2 rounded-lg font-medium transition-colors ${
              isSaving
                ? "bg-blue-400 cursor-not-allowed"
                : "bg-blue-600 hover:bg-blue-700"
            } text-white`}
          >
            {isSaving ? "⏳ Menyimpan..." : "💾 Simpan Data Sekolah"}
          </button>
        </div>
      </div>
    </div>
  );
};

const StudentDataTab: React.FC<{
  students: Student[];
  onRefresh: () => void;
  uniqueClasses: string[];
}> = ({ students, onRefresh, uniqueClasses }) => {
  const [nisn, setNisn] = useState("");
  const [nama, setNama] = useState("");
  const [kelas, setKelas] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");

  // State untuk bulk import
  const [showBulkImport, setShowBulkImport] = useState(false);
  const [bulkNisn, setBulkNisn] = useState("");
  const [bulkNama, setBulkNama] = useState("");
  const [bulkKelas, setBulkKelas] = useState("");

  // State untuk loading
  const [isSaving, setIsSaving] = useState(false);
  const [isBulkSaving, setIsBulkSaving] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);

  const handleSubmit = () => {
    if (!nisn || !nama || !kelas) {
      alert("⚠️ Semua field wajib diisi!");
      return;
    }

    setIsSaving(true);

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "siswa",
        nisn,
        nama,
        kelas,
      }),
    })
      .then(() => {
        alert("✅ Siswa berhasil ditambahkan!");
        setNisn("");
        setNama("");
        setKelas("");
        onRefresh();
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal menambahkan siswa.");
        setIsSaving(false);
      });
  };

  const handleBulkImport = () => {
    // Validasi input
    if (!bulkNisn.trim() || !bulkNama.trim() || !bulkKelas.trim()) {
      alert("⚠️ Semua field data massal wajib diisi!");
      return;
    }

    // Parse data dari textarea
    const nisnLines = bulkNisn
      .trim()
      .split("\n")
      .filter((line) => line.trim());
    const namaLines = bulkNama
      .trim()
      .split("\n")
      .filter((line) => line.trim());
    const kelasLines = bulkKelas
      .trim()
      .split("\n")
      .filter((line) => line.trim());

    // Validasi jumlah baris harus sama
    if (
      nisnLines.length !== namaLines.length ||
      namaLines.length !== kelasLines.length
    ) {
      alert("⚠️ Jumlah baris data NISN, Nama, dan Kelas harus sama!");
      return;
    }

    if (nisnLines.length === 0) {
      alert("⚠️ Tidak ada data yang valid untuk diimport!");
      return;
    }

    // Konfirmasi sebelum import
    if (!confirm(`Akan menambahkan ${nisnLines.length} siswa. Lanjutkan?`)) {
      return;
    }

    setIsBulkSaving(true);

    // Prepare data untuk bulk import
    const students = nisnLines.map((nisn, index) => ({
      nisn: nisn.trim(),
      nama: namaLines[index].trim(),
      kelas: kelasLines[index].trim(),
    }));

    // Kirim dalam satu request
    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "bulk_siswa",
        students: students,
      }),
    })
      .then(() => {
        // Karena mode no-cors, kita tidak bisa membaca response
        // Jadi kita anggap berhasil dan biarkan user refresh manual jika diperlukan
        alert(
          `✅ Data massal berhasil dikirim! Total: ${students.length} siswa`
        );

        // Reset form dan refresh data
        setBulkNisn("");
        setBulkNama("");
        setBulkKelas("");
        setShowBulkImport(false);
        onRefresh();
        setIsBulkSaving(false);
      })
      .catch((error) => {
        console.error("Error:", error);
        alert(
          "❌ Terjadi kesalahan saat import data massal. Pastikan:\n1. URL endpoint sudah benar\n2. Google Apps Script sudah di-deploy\n3. Koneksi internet stabil"
        );
        setIsBulkSaving(false);
      });
  };

  const handleEditStudent = (student: Student) => {
    const newNisn = prompt("Edit NISN:", student.nisn ?? undefined);
    const newName = prompt("Edit nama siswa:", student.name ?? undefined);
    const newClass = prompt("Edit kelas siswa:", student.kelas ?? undefined);

    if (newNisn && newName && newClass) {
      setIsEditing(true);

      fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "edit",
          nisnLama: student.nisn,
          nisnBaru: newNisn,
          nama: newName,
          kelas: newClass,
        }),
      })
        .then(() => {
          alert("✅ Data siswa berhasil diperbarui");
          onRefresh();
          setIsEditing(false);
        })
        .catch(() => {
          alert("❌ Gagal memperbarui data");
          setIsEditing(false);
        });
    }
  };

  const handleDeleteStudent = (nisn: string | null | undefined) => {
    if (!nisn) {
      alert("❌ NISN tidak valid untuk penghapusan.");
      return;
    }
    if (confirm("Yakin ingin menghapus siswa ini?")) {
      setIsDeleting(true);

      fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "delete",
          nisn: nisn,
        }),
      })
        .then(() => {
          alert("🗑️ Data siswa berhasil dihapus");
          onRefresh();
          setIsDeleting(false);
        })
        .catch(() => {
          alert("❌ Gagal menghapus siswa");
          setIsDeleting(false);
        });
    }
  };

  const filteredStudents = React.useMemo(() => {
    if (!searchQuery.trim() && selectedKelas === "Semua") return students;
    const query = searchQuery.toLowerCase().trim();
    return students.filter((student) => {
      const matchesSearchQuery =
        !searchQuery.trim() ||
        (student.name && String(student.name).toLowerCase().includes(query)) ||
        (student.nisn && String(student.nisn).toLowerCase().includes(query));
      const matchesKelas =
        selectedKelas === "Semua" ||
        (student.kelas && String(student.kelas).trim() === selectedKelas);
      return matchesSearchQuery && matchesKelas;
    });
  }, [students, searchQuery, selectedKelas]);

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      {/* Form Tambah Siswa Tunggal */}
      <div className="bg-white p-6 rounded-lg shadow-md mb-6">
        <h2 className="text-xl font-bold mb-4 text-center text-blue-600">
          Tambah Data Siswa
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
          <input
            type="text"
            placeholder="NISN"
            value={nisn}
            onChange={(e) => setNisn(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          />
          <input
            type="text"
            placeholder="Nama Siswa"
            value={nama}
            onChange={(e) => setNama(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          />
          <input
            type="text"
            placeholder="Kelas"
            value={kelas}
            onChange={(e) => setKelas(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          />
        </div>
        <div className="flex justify-center gap-4">
          <button
            onClick={handleSubmit}
            disabled={isSaving}
            className={`px-6 py-2 rounded-lg font-medium transition-colors ${
              isSaving
                ? "bg-blue-400 cursor-not-allowed"
                : "bg-blue-600 hover:bg-blue-700"
            } text-white`}
          >
            {isSaving ? "⏳ Menyimpan..." : "➕ Tambah Siswa"}
          </button>
          <button
            onClick={() => setShowBulkImport(!showBulkImport)}
            disabled={isSaving}
            className={`px-6 py-2 rounded-lg font-medium transition-colors ${
              isSaving
                ? "bg-green-400 cursor-not-allowed"
                : "bg-green-600 hover:bg-green-700"
            } text-white`}
          >
            📋 Tambah Data Massal
          </button>
        </div>
      </div>

      {/* Form Bulk Import */}
      {showBulkImport && (
        <div className="bg-white p-6 rounded-lg shadow-md mb-6 border-2 border-green-200">
          <h2 className="text-xl font-bold mb-4 text-center text-green-600">
            Import Data Massal
          </h2>
          <div className="mb-4 p-4 bg-green-50 rounded-lg">
            <p className="text-sm text-green-700 mb-2">
              <strong>Cara penggunaan:</strong>
            </p>
            <p className="text-sm text-green-600">
              1. Copy data dari Excel (pilih kolom NISN, Nama, dan Kelas secara
              terpisah)
              <br />
              2. Paste ke masing-masing kotak di bawah ini
              <br />
              3. Pastikan jumlah baris di setiap kolom sama
              <br />
              4. Klik "Kirim Data Massal"
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                NISN (satu per baris)
              </label>
              <textarea
                placeholder="34534534&#10;56565656&#10;12345678"
                value={bulkNisn}
                onChange={(e) => setBulkNisn(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Nama (satu per baris)
              </label>
              <textarea
                placeholder="Andika&#10;Alisa&#10;Budi"
                value={bulkNama}
                onChange={(e) => setBulkNama(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Kelas (satu per baris)
              </label>
              <textarea
                placeholder="3&#10;4&#10;5"
                value={bulkKelas}
                onChange={(e) => setBulkKelas(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
          </div>

          <div className="flex justify-center gap-4">
            <button
              onClick={handleBulkImport}
              disabled={isBulkSaving}
              className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                isBulkSaving
                  ? "bg-green-400 cursor-not-allowed"
                  : "bg-green-600 hover:bg-green-700"
              } text-white`}
            >
              {isBulkSaving ? "⏳ Menyimpan..." : "📤 Kirim Data Massal"}
            </button>
            <button
              onClick={() => {
                setBulkNisn("");
                setBulkNama("");
                setBulkKelas("");
                setShowBulkImport(false);
              }}
              disabled={isBulkSaving}
              className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                isBulkSaving
                  ? "bg-gray-400 cursor-not-allowed"
                  : "bg-gray-500 hover:bg-gray-600"
              } text-white`}
            >
              ❌ Batal
            </button>
          </div>
        </div>
      )}

      {/* Pencarian Siswa */}
      <div className="bg-white p-6 rounded-lg shadow-md mb-6">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Pencarian Siswa
        </h3>
        <div className="mb-4">
          <input
            type="text"
            placeholder="Cari berdasarkan nama atau NISN..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
          />
        </div>
        <div className="mb-4">
          <select
            value={selectedKelas}
            onChange={(e) => setSelectedKelas(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-4 py-2 shadow-sm bg-white"
          >
            {uniqueClasses.map((kelas) => (
              <option key={kelas} value={kelas}>
                {kelas}
              </option>
            ))}
          </select>
        </div>
      </div>

      {/* Daftar Siswa */}
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Daftar Siswa ({filteredStudents.length})
        </h3>
        {filteredStudents.length === 0 ? (
          <p className="text-center text-gray-500 py-8">
            {searchQuery || selectedKelas !== "Semua"
              ? "Tidak ada siswa yang cocok dengan pencarian atau filter kelas."
              : "Belum ada data siswa."}
          </p>
        ) : (
          <div className="space-y-3">
            {filteredStudents.map((s) => (
              <div
                key={s.id}
                className="flex justify-between items-center bg-gray-50 border border-gray-200 px-4 py-3 rounded-lg"
              >
                <div>
                  <p className="font-medium text-gray-800">{s.name || "N/A"}</p>
                  <p className="text-sm text-gray-600">
                    NISN: {s.nisn || "N/A"} | Kelas: {s.kelas || "N/A"}
                  </p>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => handleEditStudent(s)}
                    disabled={isEditing}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isEditing
                        ? "bg-yellow-400 cursor-not-allowed"
                        : "bg-yellow-500 hover:bg-yellow-600"
                    } text-white`}
                  >
                    {isEditing ? "⏳" : "✏️"} Edit
                  </button>
                  <button
                    onClick={() => handleDeleteStudent(s.nisn)}
                    disabled={isDeleting}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isDeleting
                        ? "bg-red-400 cursor-not-allowed"
                        : "bg-red-500 hover:bg-red-600"
                    } text-white`}
                  >
                    {isDeleting ? "⏳" : "🗑️"} Hapus
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

const AttendanceTab: React.FC<{
  students: Student[];
  onRecapRefresh: () => void;
}> = ({ students, onRecapRefresh }) => {
  const [attendance, setAttendance] = useState<AttendanceRecord>({});
  const [date, setDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [showDebugInfo, setShowDebugInfo] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false); // Add this line

  const uniqueClasses = React.useMemo(() => {
    console.log("Memproses siswa untuk kelas:", students);

    const classSet = new Set<string>();

    students.forEach((student) => {
      console.log(
        "Siswa:",
        student.name,
        "Kelas:",
        student.kelas,
        "Tipe:",
        typeof student.kelas
      );

      let kelasValue = student.kelas;

      if (kelasValue != null) {
        kelasValue = String(kelasValue).trim();

        if (
          kelasValue !== "" &&
          kelasValue !== "undefined" &&
          kelasValue !== "null"
        ) {
          classSet.add(kelasValue);
        }
      }
    });

    const classes = Array.from(classSet).sort((a, b) => {
      const aIsNum = /^\d+$/.test(a);
      const bIsNum = /^\d+$/.test(b);

      if (aIsNum && bIsNum) {
        return parseInt(a) - parseInt(b);
      } else if (aIsNum && !bIsNum) {
        return -1;
      } else if (!aIsNum && bIsNum) {
        return 1;
      } else {
        return a.localeCompare(b);
      }
    });

    console.log("Kelas unik yang ditemukan:", classes);
    return ["Semua", ...classes];
  }, [students]);

  const filteredStudents = React.useMemo(() => {
    if (selectedKelas === "Semua") {
      return students;
    }

    return students.filter((student) => {
      if (student.kelas == null) return false;
      const studentKelas = String(student.kelas).trim();
      const result = studentKelas === selectedKelas;
      console.log(
        `Menyaring: ${student.name} (${studentKelas}) === ${selectedKelas} = ${result}`
      );
      return result;
    });
  }, [students, selectedKelas]);

  useEffect(() => {
    if (students.length && !attendance[date]) {
      const init: { [key: string]: AttendanceStatus } = {};
      students.forEach((s) => (init[s.id] = "Hadir"));
      setAttendance((prev) => ({ ...prev, [date]: init }));
    }
  }, [date, students, attendance]);

  const setStatus = (sid: string, status: AttendanceStatus) => {
    setAttendance((prev) => ({
      ...prev,
      [date]: { ...prev[date], [sid]: status },
    }));
  };

  const handleSave = () => {
    setIsSaving(true); // Set saving state to true

    const formattedDate = formatDateDDMMYYYY(date);
    const studentsToSave =
      selectedKelas === "Semua" ? students : filteredStudents;

    const data = studentsToSave.map((s) => ({
      tanggal: formattedDate,
      nama: s.name || "N/A",
      kelas: s.kelas || "N/A",
      nisn: s.nisn || "N/A",
      status: attendance[date]?.[s.id] || "Hadir",
    }));

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    })
      .then(() => {
        const message =
          selectedKelas === "Semua"
            ? "✅ Data absensi semua kelas berhasil dikirim!"
            : `✅ Data absensi kelas ${selectedKelas} berhasil dikirim!`;
        alert(message);
        onRecapRefresh();
        setIsSaving(false); // Reset saving state on success
      })
      .catch(() => {
        alert("❌ Gagal kirim data absensi.");
        setIsSaving(false); // Reset saving state on error
      });
  };

  const statusColor: Record<AttendanceStatus, string> = {
    Hadir: "bg-green-500",
    Izin: "bg-yellow-400",
    Sakit: "bg-blue-400",
    Alpha: "bg-red-500",
  };

  const getAttendanceSummary = (): StatusSummary => {
    const summary: StatusSummary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredStudents.forEach((s) => {
      const status = (attendance[date]?.[s.id] || "Hadir") as AttendanceStatus;
      summary[status]++;
    });
    return summary;
  };

  const attendanceSummary = getAttendanceSummary();

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📋 Absensi Siswa
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Tanggal</p>
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm"
            />
          </div>

          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => {
                console.log("Mengubah filter kelas ke:", e.target.value);
                setSelectedKelas(e.target.value);
              }}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>

          <div className="text-center">
            <button
              onClick={() => setShowDebugInfo(!showDebugInfo)}
              className="text-sm bg-gray-200 hover:bg-gray-300 px-1 py-0.5 rounded-lg"
            >
              🔍 Info Debug
            </button>
          </div>
        </div>

        {showDebugInfo && (
          <div className="mb-6 bg-yellow-50 border border-yellow-200 rounded-lg p-4">
            <h4 className="font-semibold text-yellow-800 mb-2">
              Informasi Debug:
            </h4>
            <div className="text-sm text-yellow-700 space-y-1">
              <p>
                <strong>Total Siswa:</strong> {students.length}
              </p>
              <p>
                <strong>Kelas yang Tersedia:</strong> {uniqueClasses.join(", ")}
              </p>
              <p>
                <strong>Kelas Terpilih:</strong> {selectedKelas}
              </p>
              <p>
                <strong>Siswa Terfilter:</strong> {filteredStudents.length}
              </p>
            </div>
            <div className="mt-3">
              <p className="font-semibold text-yellow-800 mb-1">
                Detail Data Siswa per Kelas:
              </p>
              <div className="max-h-32 overflow-y-auto text-xs">
                {uniqueClasses.slice(1).map((kelas) => {
                  const siswaKelas = students.filter(
                    (s) => String(s.kelas).trim() === kelas
                  );
                  return (
                    <div key={kelas} className="mb-1">
                      <strong>Kelas {kelas}:</strong> {siswaKelas.length} siswa
                      {siswaKelas.slice(0, 3).map((s) => (
                        <div key={s.id} className="ml-4 text-gray-600">
                          • {s.name || "N/A"} (NISN: {s.nisn || "N/A"}, Kelas:{" "}
                          {s.kelas || "N/A"})
                        </div>
                      ))}
                      {siswaKelas.length > 3 && (
                        <div className="ml-4 text-gray-500">
                          ... dan {siswaKelas.length - 3} lainnya
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
            <div className="mt-3">
              <p className="font-semibold text-yellow-800 mb-1">
                Sampel Data Siswa Mentah:
              </p>
              <div className="max-h-24 overflow-y-auto text-xs bg-white p-2 rounded border">
                {students.slice(0, 5).map((s, idx) => (
                  <div key={idx} className="text-gray-600">
                    {idx + 1}. {s.name || "N/A"} | Kelas: "{s.kelas || "N/A"}"
                    (type: {typeof s.kelas}) | NISN: "{s.nisn || "N/A"}" (type:{" "}
                    {typeof s.nisn})
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        <div className="mb-4 text-center">
          <p className="text-sm text-gray-600">
            Menampilkan: <span className="font-semibold">{selectedKelas}</span>{" "}
            • Tanggal:{" "}
            <span className="font-semibold">{formatDateDDMMYYYY(date)}</span> •
            Total Siswa:{" "}
            <span className="font-semibold">{filteredStudents.length}</span>
          </p>
        </div>

        {students.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Belum ada data siswa.</p>
            <p className="text-sm text-gray-400 mt-2">
              Silakan tambah data siswa terlebih dahulu di tab "Data Siswa"
            </p>
          </div>
        ) : filteredStudents.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada siswa di kelas {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Pilih kelas lain atau ubah filter ke "Semua"
            </p>
          </div>
        ) : (
          <>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
              <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
                <div className="text-green-600 font-bold text-lg">
                  {attendanceSummary.Hadir}
                </div>
                <div className="text-green-700 text-sm">Hadir</div>
              </div>
              <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
                <div className="text-yellow-600 font-bold text-lg">
                  {attendanceSummary.Izin}
                </div>
                <div className="text-yellow-700 text-sm">Izin</div>
              </div>
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
                <div className="text-blue-600 font-bold text-lg">
                  {attendanceSummary.Sakit}
                </div>
                <div className="text-blue-700 text-sm">Sakit</div>
              </div>
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
                <div className="text-red-600 font-bold text-lg">
                  {attendanceSummary.Alpha}
                </div>
                <div className="text-red-700 text-sm">Alpha</div>
              </div>
            </div>

            <div className="space-y-4 mb-6 overflow-x-auto">
              <table className="w-full border-collapse">
                <tbody>
                  {filteredStudents.map((s) => (
                    <tr key={s.id} className="border-b border-gray-200">
                      <td style={{ width: "6cm" }} className="p-2">
                        <p className="text-base font-semibold text-gray-800">
                          {s.name || "N/A"}
                        </p>
                        <p className="text-sm text-gray-500">
                          Kelas {s.kelas || "N/A"} • NISN: {s.nisn || "N/A"}
                        </p>
                      </td>
                      <td style={{ width: "5cm" }} className="p-2">
                        <div className="flex justify-between">
                          {(["Hadir", "Izin", "Sakit", "Alpha"] as const).map(
                            (status) => (
                              <button
                                key={status}
                                onClick={() => setStatus(s.id, status)}
                                style={{ width: "1cm" }}
                                className={`px-1 py-0.5 rounded-lg text-xs font-medium transition-colors ${
                                  attendance[date]?.[s.id] === status
                                    ? `${statusColor[status]} text-white`
                                    : "bg-white text-gray-700 border border-gray-300 hover:bg-gray-100"
                                }`}
                              >
                                {status}
                              </button>
                            )
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <button
              onClick={handleSave}
              disabled={isSaving}
              className={`w-full py-3 rounded-lg font-bold shadow-md transition-colors ${
                isSaving
                  ? "bg-blue-400 cursor-not-allowed"
                  : "bg-blue-600 hover:bg-blue-700"
              } text-white`}
            >
              {isSaving
                ? "⏳ Menyimpan..."
                : "💾 Simpan Absensi " +
                  (selectedKelas !== "Semua"
                    ? `Kelas ${selectedKelas}`
                    : "Semua Kelas")}
            </button>
          </>
        )}
      </div>
    </div>
  );
};

const MonthlyRecapTab: React.FC<{
  onRefresh: () => void;
  uniqueClasses: string[];
}> = ({ onRefresh, uniqueClasses }) => {
  const [recapData, setRecapData] = useState<MonthlyRecap[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [selectedBulan, setSelectedBulan] = useState<string>("Oktober");
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );
  const [placeName, setPlaceName] = useState<string>("Makassar");
  const [loading, setLoading] = useState<boolean>(true);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);

  const months = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ] as const;

  useEffect(() => {
    setLoading(true);
    console.log(
      "Mengambil data rekap dengan kelas:",
      selectedKelas,
      "dan bulan:",
      selectedBulan
    );
    fetch(
      `${endpoint}?action=monthlyRecap&kelas=${
        selectedKelas === "Semua" ? "" : selectedKelas
      }&bulan=${selectedBulan.toLowerCase()}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        console.log("Respons data rekap:", data);
        if (data.success) {
          setRecapData(data.data || []);
        } else {
          alert("❌ Gagal memuat data rekap: " + data.message);
          setRecapData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data rekap. Cek console untuk detail.");
        setRecapData([]);
        setLoading(false);
      });

    // Fetch school data
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, [selectedKelas, selectedBulan, onRefresh]);

  const filteredRecapData = React.useMemo(() => {
    if (selectedKelas === "Semua") {
      return recapData;
    }
    console.log("Menyaring data rekap untuk kelas:", selectedKelas);
    return recapData.filter((item) => {
      const itemKelas = String(item.kelas).trim();
      const result = itemKelas === selectedKelas;
      console.log("Kelas item:", itemKelas, "cocok?", result);
      return result;
    });
  }, [recapData, selectedKelas]);

  const getStatusSummary = (): StatusSummary => {
    const summary: StatusSummary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredRecapData.forEach((item) => {
      summary.Hadir += item.hadir || 0;
      summary.Alpha += item.alpa || 0;
      summary.Izin += item.izin || 0;
      summary.Sakit += item.sakit || 0;
    });
    return summary;
  };

  const statusSummary = getStatusSummary();

  const downloadExcel = () => {
    const headers = [
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const data = [
      headers,
      ...filteredRecapData.map((item) => [
        item.nama || "N/A",
        item.kelas || "N/A",
        item.hadir || 0,
        item.alpa || 0,
        item.izin || 0,
        item.sakit || 0,
        item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
      ]),
      [
        "TOTAL",
        "",
        statusSummary.Hadir,
        statusSummary.Alpha,
        statusSummary.Izin,
        statusSummary.Sakit,
        "",
      ],
      [
        "PERSEN",
        "",
        `${(
          (statusSummary.Hadir /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Alpha /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Izin /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Sakit /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        "",
      ],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = headers.map(() => ({ wch: 15 }));
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FFFF00" } },
      alignment: { horizontal: "center" },
    };
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    const percentStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    headers.forEach((header, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cellAddress] = { ...ws[cellAddress], s: headerStyle };
    });
    const totalRow = filteredRecapData.length + 1;
    ["A", "B", "C", "D", "E", "F", "G"].forEach((col, idx) => {
      const cellAddress = `${col}${totalRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: totalStyle };
    });
    const percentRow = filteredRecapData.length + 2;
    ["A", "B", "C", "D", "E", "F", "G"].forEach((col, idx) => {
      const cellAddress = `${col}${percentRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: percentStyle };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Bulanan");

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Bulanan_${selectedBulan}_${selectedKelas}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const downloadPDF = () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    // Set font to Times for consistency with typical formal documents
    doc.setFont("Times", "roman");

    // Title
    const title = `REKAP ABSENSI SISWA KELAS ${selectedKelas} ${selectedBulan.toUpperCase()} 2024`;
    doc.setFontSize(14);
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });
    currentY += 10;

    // Table headers and data
    const headers = [
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const body = filteredRecapData.map((item) => [
      item.nama || "N/A",
      item.kelas || "N/A",
      item.hadir || 0,
      item.alpa || 0,
      item.izin || 0,
      item.sakit || 0,
      item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
    ]);

    const totalRow = [
      "TOTAL",
      "",
      statusSummary.Hadir,
      statusSummary.Alpha,
      statusSummary.Izin,
      statusSummary.Sakit,
      "",
    ];

    const percentRow = [
      "PERSEN",
      "",
      `${(
        (statusSummary.Hadir /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Alpha /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Izin /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Sakit /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      "",
    ];

    autoTable(doc, {
      head: [headers],
      body: [...body, totalRow, percentRow],
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      columnStyles: {
        0: { cellWidth: 50 },
        1: { cellWidth: 20 },
        2: { cellWidth: 20 },
        3: { cellWidth: 20 },
        4: { cellWidth: 20 },
        5: { cellWidth: 20 },
        6: { cellWidth: 20 },
      },
    });

    // Update currentY after the table
    currentY = (doc as any).lastAutoTable.finalY + 10;

    // Add school data (Principal and Teacher details)
    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      // Add place and date above Guru Kelas, centered
      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${placeName}, ${formattedDate}`;
      const rightColumnX = pageWidth - margin - 50; // Signature width is 50
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5; // Keep close to "Guru Kelas"

      // Principal Section
      const principalText = [
        "Kepala Sekolah,",
        "",
        "",
        `( ${schoolData.namaKepsek || "N/A"} )`,
        `NIP: ${schoolData.nipKepsek || "N/A"}`,
      ];
      const teacherText = [
        "Guru Kelas,",
        "",
        "",
        `( ${schoolData.namaGuru || "N/A"} )`,
        `NIP: ${schoolData.nipGuru || "N/A"}`,
      ];

      // Calculate width for signatures
      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      // Principal signature and text
      if (schoolData.ttdKepsek) {
        doc.addImage(
          schoolData.ttdKepsek,
          "PNG",
          leftColumnX + 10,
          currentY - 3,
          signatureWidth,
          signatureHeight
        );
      }

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama kepala sekolah dengan format bold dan underline
      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to principal name
      const principalNameText = principalName;
      const textWidth = doc.getTextWidth(principalNameText);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );

      // Teacher signature and text
      if (schoolData.ttdGuru) {
        doc.addImage(
          schoolData.ttdGuru,
          "PNG",
          rightColumnX + 10,
          currentY - 5,
          signatureWidth,
          signatureHeight
        );
      }

      // Pisahkan "Guru Kelas" dengan posisi yang lebih tinggi
      doc.text("Guru Kelas,", rightColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama guru dengan format bold dan underline
      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to teacher name
      const teacherNameText = teacherName;
      const teacherTextWidth = doc.getTextWidth(teacherNameText);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Bulanan_${selectedBulan}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📊 Rekap Absensi Bulanan
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => {
                console.log("Mengubah filter kelas ke:", e.target.value);
                setSelectedKelas(e.target.value);
              }}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Pilih Bulan</p>
            <select
              value={selectedBulan}
              onChange={(e) => setSelectedBulan(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {months.map((month) => (
                <option key={month} value={month}>
                  {month}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Rekap Bulanan pada
            File PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Nama Tempat</p>
              <input
                type="text"
                value={placeName}
                onChange={(e) => setPlaceName(e.target.value)}
                placeholder="Masukkan nama tempat"
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
            <div className="text-green-600 font-bold text-lg">
              {statusSummary.Hadir}
            </div>
            <div className="text-green-700 text-sm">Hadir</div>
          </div>
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
            <div className="text-yellow-600 font-bold text-lg">
              {statusSummary.Izin}
            </div>
            <div className="text-yellow-700 text-sm">Izin</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
            <div className="text-blue-600 font-bold text-lg">
              {statusSummary.Sakit}
            </div>
            <div className="text-blue-700 text-sm">Sakit</div>
          </div>
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
            <div className="text-red-600 font-bold text-lg">
              {statusSummary.Alpha}
            </div>
            <div className="text-red-700 text-sm">Alpha</div>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat rekap...</p>
          </div>
        ) : filteredRecapData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data rekap untuk {selectedBulan} kelas {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Coba pilih kelas atau bulan lain.
            </p>
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Nama
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Kelas
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Hadir
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Alpha
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Izin
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Sakit
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      % Hadir
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecapData.map((item, index) => (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.hadir || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.alpa || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.izin || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.sakit || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.persenHadir !== undefined
                          ? `${item.persenHadir}%`
                          : "N/A"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 flex gap-4 justify-center">
              <button
                onClick={downloadExcel}
                className="px-1 py-0.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download Excel
              </button>
              <button
                onClick={downloadPDF}
                className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📄 Download PDF
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const GraphTab: React.FC<{
  uniqueClasses: string[];
}> = ({ uniqueClasses }) => {
  const [graphData, setGraphData] = useState<GraphData>({
    Januari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Februari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Maret: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    April: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Mei: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Juni: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Juli: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Agustus: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    September: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Oktober: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    November: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Desember: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
  });
  const [selectedKelas, setSelectedKelas] = useState<string>(
    uniqueClasses.length > 0 ? uniqueClasses[0] : "Tidak Ada"
  );
  const [selectedSemester, setSelectedSemester] = useState<"1" | "2">("2");
  const [statusVisibility, setStatusVisibility] = useState<StatusVisibility>({
    Hadir: true,
    Alpha: true,
    Izin: true,
    Sakit: true,
  });
  const [loading, setLoading] = useState<boolean>(true);

  const uniqueClassesWithDefault = React.useMemo(() => {
    return ["Tidak Ada", ...uniqueClasses.filter((kelas) => kelas !== "Semua")];
  }, [uniqueClasses]);

  useEffect(() => {
    setLoading(true);
    fetch(
      `${endpoint}?action=graphData&kelas=${
        selectedKelas === "Tidak Ada" ? "" : selectedKelas
      }&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setGraphData(data.data || {});
        } else {
          alert("❌ Gagal memuat data grafik: " + data.message);
          setGraphData({
            Januari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Februari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Maret: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            April: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Mei: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Juni: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Juli: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Agustus: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            September: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Oktober: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            November: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Desember: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
          });
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data grafik. Cek console untuk detail.");
        setLoading(false);
      });
  }, [selectedKelas, selectedSemester]);

  const semesterMonths: Record<"1" | "2", string[]> = {
    "1": ["Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    "2": ["Januari", "Februari", "Maret", "April", "Mei", "Juni"],
  };

  const chartData: ChartData<"bar", number[], string> = {
    labels: semesterMonths[selectedSemester],
    datasets: [
      ...(statusVisibility.Hadir
        ? [
            {
              label: "Hadir",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Hadir || 0
              ),
              backgroundColor: "rgba(75, 192, 192, 0.6)",
              borderColor: "rgba(75, 192, 192, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Alpha
        ? [
            {
              label: "Alpha",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Alpha || 0
              ),
              backgroundColor: "rgba(255, 99, 132, 0.6)",
              borderColor: "rgba(255, 99, 132, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Izin
        ? [
            {
              label: "Izin",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Izin || 0
              ),
              backgroundColor: "rgba(255, 205, 86, 0.6)",
              borderColor: "rgba(255, 205, 86, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Sakit
        ? [
            {
              label: "Sakit",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Sakit || 0
              ),
              backgroundColor: "rgba(54, 162, 235, 0.6)",
              borderColor: "rgba(54, 162, 235, 1)",
              borderWidth: 1,
            },
          ]
        : []),
    ],
  };

  const chartOptions: ChartOptions<"bar"> = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: "top" as const,
        onClick: (
          e: ChartEvent,
          legendItem: LegendItem,
          legend: {
            chart: {
              data: { datasets: { hidden?: boolean }[] };
              update: () => void;
            };
          }
        ) => {
          const index = legendItem.datasetIndex;
          if (index !== undefined) {
            const ci = legend.chart.data.datasets[index];
            ci.hidden = !ci.hidden;
            legend.chart.update();
            setStatusVisibility((prev) => ({
              ...prev,
              [legendItem.text as keyof StatusVisibility]: !ci.hidden,
            }));
          }
        },
      },
      title: {
        display: true,
        text: `Persentase Kehadiran Kelas ${selectedKelas} Semester ${selectedSemester} 2025`,
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        max: 100,
        ticks: {
          stepSize: 10,
          font: {
            size: 10,
          },
          autoSkip: false,
          maxTicksLimit: 11,
        },
        title: { display: true, text: "Persentase (%)" },
      },
      x: {
        ticks: {
          font: {
            size: 10,
          },
        },
      },
    },
  };

  const handleStatusToggle = (status: keyof StatusVisibility) => {
    setStatusVisibility((prev) => ({
      ...prev,
      [status]: !prev[status],
    }));
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📈 Grafik Kehadiran
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClassesWithDefault.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Semester</p>
            <select
              value={selectedSemester}
              onChange={(e) => setSelectedSemester(e.target.value as "1" | "2")}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              <option value="1">Semester 1 (Juli-Des)</option>
              <option value="2">Semester 2 (Jan-Jun)</option>
            </select>
          </div>
        </div>

        <div className="mb-6 flex flex-wrap gap-4 justify-center">
          {(["Hadir", "Alpha", "Izin", "Sakit"] as const).map((status) => (
            <label key={status} className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={statusVisibility[status]}
                onChange={() => handleStatusToggle(status)}
                className="h-4 w-4 text-blue-600 rounded"
              />
              <span className="text-sm text-gray-700">{status}</span>
            </label>
          ))}
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat grafik...</p>
          </div>
        ) : selectedKelas === "Tidak Ada" ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Tidak ada data untuk ditampilkan.</p>
          </div>
        ) : (
          <div
            className="h-96"
            style={{
              minHeight: "300px",
              maxHeight: "500px",
            }}
          >
            <Bar data={chartData} options={chartOptions} />
          </div>
        )}
      </div>
    </div>
  );
};

const AttendanceHistoryTab: React.FC<{
  students: Student[];
  uniqueClasses: string[];
  onRefresh: () => void;
}> = ({ students, uniqueClasses, onRefresh }) => {
  const [attendanceData, setAttendanceData] = useState<AttendanceHistory[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [selectedNama, setSelectedNama] = useState<string>("Semua");
  const [selectedDate, setSelectedDate] = useState<string>("");
  const [editedRecords, setEditedRecords] = useState<
    Record<string, AttendanceStatus>
  >({});
  const [isDeleting, setIsDeleting] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);

  const formatDateToDDMMYYYY = (dateStr: string): string => {
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) return dateStr;
    const [datePart] = dateStr.split("T");
    const [year, month, day] = datePart.split("-");
    return `${day}/${month}/${year}`;
  };

  const filterNonFormulaData = (
    data: AttendanceHistory[]
  ): AttendanceHistory[] => {
    return data.filter((record) => {
      const hasValidTanggal =
        record.tanggal &&
        !record.tanggal.toString().startsWith("=") &&
        record.tanggal.toString().trim() !== "" &&
        record.tanggal.toString() !== "#N/A" &&
        record.tanggal.toString() !== "#REF!" &&
        record.tanggal.toString() !== "#VALUE!" &&
        record.tanggal.toString() !== "#ERROR!" &&
        !record.tanggal.toString().includes("FORMULA");

      const hasValidNama =
        record.nama &&
        !record.nama.toString().startsWith("=") &&
        record.nama.toString().trim() !== "" &&
        record.nama.toString() !== "#N/A" &&
        record.nama.toString() !== "#REF!" &&
        record.nama.toString() !== "#VALUE!" &&
        record.nama.toString() !== "#ERROR!" &&
        !record.nama.toString().includes("FORMULA");

      const hasValidNisn =
        record.nisn &&
        !record.nisn.toString().startsWith("=") &&
        record.nisn.toString().trim() !== "" &&
        record.nisn.toString() !== "#N/A" &&
        record.nisn.toString() !== "#REF!" &&
        record.nisn.toString() !== "#VALUE!" &&
        record.nisn.toString() !== "#ERROR!" &&
        !record.nisn.toString().includes("FORMULA");

      const hasValidKelas =
        record.kelas &&
        !record.kelas.toString().startsWith("=") &&
        record.kelas.toString().trim() !== "" &&
        record.kelas.toString() !== "#N/A" &&
        record.kelas.toString() !== "#REF!" &&
        record.kelas.toString() !== "#VALUE!" &&
        record.kelas.toString() !== "#ERROR!" &&
        !record.kelas.toString().includes("FORMULA");

      const hasValidStatus =
        record.status &&
        !record.status.toString().startsWith("=") &&
        record.status.toString().trim() !== "" &&
        record.status.toString() !== "#N/A" &&
        record.status.toString() !== "#REF!" &&
        record.status.toString() !== "#VALUE!" &&
        record.status.toString() !== "#ERROR!" &&
        !record.status.toString().includes("FORMULA") &&
        ["Hadir", "Izin", "Sakit", "Alpha"].includes(record.status.toString());

      // Tambahkan validasi untuk kolom duplikat: hanya ambil jika nilai eksak "-", dan bukan formula/error
      const hasValidDuplikat =
        record.duplikat &&
        record.duplikat.toString().trim() === "-" && // Hanya ambil jika "-"
        !record.duplikat.toString().startsWith("=") &&
        record.duplikat.toString() !== "#N/A" &&
        record.duplikat.toString() !== "#REF!" &&
        record.duplikat.toString() !== "#VALUE!" &&
        record.duplikat.toString() !== "#ERROR!" &&
        !record.duplikat.toString().includes("FORMULA");

      return (
        hasValidTanggal &&
        hasValidNama &&
        hasValidNisn &&
        hasValidKelas &&
        hasValidStatus &&
        hasValidDuplikat // Tambahkan ini untuk memfilter berdasarkan duplikat = "-"
      );
    });
  };

  useEffect(() => {
    setLoading(true);
    fetch(`${endpoint}?action=attendanceHistory`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          const newData = data.data || [];
          const filteredData = filterNonFormulaData(newData);

          const updatedData = filteredData.map((record: AttendanceHistory) => {
            const key = `${record.tanggal}_${record.nisn}`;
            return {
              ...record,
              status: editedRecords[key] || record.status,
            };
          });
          setAttendanceData(updatedData);
        } else {
          if (data.message === "Tidak ada data di sheet Absensi") {
            setAttendanceData([]);
          } else {
            alert("❌ Gagal memuat data riwayat absensi: " + data.message);
            setAttendanceData([]);
          }
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert(
          "❌ Gagal memuat data riwayat absensi. Cek console untuk detail."
        );
        setLoading(false);
      });
  }, [onRefresh]);

  const handleUpdateStatus = (
    record: AttendanceHistory,
    newStatus: AttendanceStatus
  ) => {
    const key = `${record.tanggal}_${record.nisn}`;
    setEditedRecords((prev) => ({
      ...prev,
      [key]: newStatus,
    }));
    setAttendanceData((prev) =>
      prev.map((item) =>
        item.tanggal === record.tanggal && item.nisn === record.nisn
          ? { ...item, status: newStatus }
          : item
      )
    );
  };

  const handleDeleteAllAttendance = () => {
    if (
      confirm(
        "Yakin ingin menghapus semua data absensi di sheet 'absensi'? Header tidak akan terhapus."
      )
    ) {
      setIsDeleting(true);
      fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "deleteAllAttendance",
          sheetName: "absensi",
        }),
      })
        .then(() => {
          alert(
            "✅ Semua data absensi di sheet 'absensi' berhasil dihapus. Header tetap utuh."
          );
          setAttendanceData([]);
          setEditedRecords({});
          onRefresh();
        })
        .catch(() =>
          alert("❌ Gagal menghapus data absensi di sheet 'absensi'.")
        )
        .finally(() => setIsDeleting(false));
    }
  };

  const filteredStudents = React.useMemo(() => {
    return selectedKelas === "Semua"
      ? students
      : students.filter(
          (student) => String(student.kelas).trim() === selectedKelas
        );
  }, [students, selectedKelas]);

  const uniqueNames = React.useMemo(() => {
    const names = filteredStudents
      .map((student) => student.name)
      .filter((name): name is string => name != null && name.trim() !== "");
    return ["Semua", ...Array.from(new Set(names)).sort()];
  }, [filteredStudents]);

  const filteredAttendanceData = React.useMemo(() => {
    return attendanceData.filter((record) => {
      const matchesKelas =
        selectedKelas === "Semua" ||
        String(record.kelas).trim() === selectedKelas;
      const matchesNama =
        selectedNama === "Semua" || record.nama === selectedNama;
      const matchesDate =
        selectedDate === "" || record.tanggal === selectedDate;
      return matchesKelas && matchesNama && matchesDate;
    });
  }, [attendanceData, selectedKelas, selectedNama, selectedDate]);

  useEffect(() => {
    setSelectedNama("Semua");
  }, [selectedKelas]);

  const handleDateChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const isoDate = e.target.value;
    if (isoDate) {
      setSelectedDate(formatDateToDDMMYYYY(isoDate));
    } else {
      setSelectedDate("");
    }
  };

  const statusColor: Record<AttendanceStatus, string> = {
    Hadir: "bg-green-500",
    Izin: "bg-yellow-400",
    Sakit: "bg-blue-400",
    Alpha: "bg-red-500",
  };

  const saveChanges = () => {
    if (Object.keys(editedRecords).length === 0) {
      alert("⚠️ Tidak ada perubahan yang perlu disimpan.");
      return;
    }

    setIsSaving(true);
    const updates = Object.entries(editedRecords).map(([key, status]) => {
      const [tanggal, nisn] = key.split("_");
      return { tanggal, nisn, status };
    });

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "bulkUpdateAttendance",
        updates,
      }),
    })
      .then(() => {
        alert("✅ Perubahan status kehadiran berhasil disimpan!");
        setEditedRecords({});
        onRefresh();
      })
      .catch(() => alert("❌ Gagal menyimpan perubahan status kehadiran."))
      .finally(() => setIsSaving(false));
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📜 Riwayat Absensi
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Nama Siswa</p>
            <select
              value={selectedNama}
              onChange={(e) => setSelectedNama(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueNames.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Tanggal</p>
            <input
              type="date"
              onChange={handleDateChange}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm"
            />
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat riwayat...</p>
          </div>
        ) : filteredAttendanceData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data riwayat absensi yang sesuai dengan filter.
            </p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="min-w-full border-collapse border border-gray-200">
              <thead>
                <tr className="bg-gray-100">
                  <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                    Tanggal
                  </th>
                  <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                    Nama
                  </th>
                  <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                    Kelas
                  </th>
                  <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                    NISN
                  </th>
                  <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                    Status
                  </th>
                </tr>
              </thead>
              <tbody>
                {filteredAttendanceData.map((record, index) => {
                  const key = `${record.tanggal}_${record.nisn}`;
                  const currentStatus = editedRecords[key] || record.status;
                  return (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {record.tanggal || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {record.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {record.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {record.nisn || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center">
                        <div className="flex justify-between">
                          {(["Hadir", "Izin", "Sakit", "Alpha"] as const).map(
                            (status) => (
                              <button
                                key={status}
                                onClick={() =>
                                  handleUpdateStatus(record, status)
                                }
                                style={{ width: "2cm" }}
                                className={`px-1 py-0.5 rounded-lg text-xs font-medium transition-colors ${
                                  currentStatus === status
                                    ? `${statusColor[status]} text-white`
                                    : "bg-white text-gray-700 border border-gray-300 hover:bg-gray-100"
                                }`}
                              >
                                {status}
                              </button>
                            )
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        <div className="mt-6 flex flex-col md:flex-row gap-4 justify-center">
          {Object.keys(editedRecords).length > 0 && (
            <button
              onClick={saveChanges}
              disabled={isSaving}
              className={`px-6 py-2 text-white rounded-lg font-medium ${
                isSaving
                  ? "bg-blue-400 cursor-not-allowed"
                  : "bg-blue-600 hover:bg-blue-700"
              }`}
            >
              {isSaving ? "Memproses..." : "💾 Simpan Perubahan"}
            </button>
          )}
          <button
            onClick={handleDeleteAllAttendance}
            disabled={isDeleting}
            className={`px-6 py-2 text-white rounded-lg font-medium ${
              isDeleting
                ? "bg-red-400 cursor-not-allowed"
                : "bg-red-600 hover:bg-red-700"
            }`}
          >
            {isDeleting ? "Memproses..." : "🗑️ Hapus Semua Data Absensi"}
          </button>
        </div>
      </div>
    </div>
  );
};

const SemesterRecapTab: React.FC<{ uniqueClasses: string[] }> = ({
  uniqueClasses,
}) => {
  const [recapData, setRecapData] = useState<SemesterRecap[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [selectedSemester, setSelectedSemester] = useState<"1" | "2">("1");
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );
  const [placeName, setPlaceName] = useState<string>("Makassar");
  const [loading, setLoading] = useState<boolean>(true);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);

  useEffect(() => {
    setLoading(true);
    const sheetName =
      selectedSemester === "1" ? SHEET_SEMESTER1 : SHEET_SEMESTER2;
    fetch(
      `${endpoint}?action=semesterRecap&kelas=${
        selectedKelas === "Semua" ? "" : selectedKelas
      }&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setRecapData(data.data || []);
        } else {
          alert(
            `❌ Gagal memuat data rekap ${
              selectedSemester === "1" ? "Semester 1" : "Semester 2"
            }: ${data.message}`
          );
          setRecapData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert(
          `❌ Gagal memuat data rekap ${
            selectedSemester === "1" ? "Semester 1" : "Semester 2"
          }. Cek console untuk detail.`
        );
        setRecapData([]);
        setLoading(false);
      });

    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, [selectedKelas, selectedSemester]);

  const filteredRecapData = React.useMemo(() => {
    if (selectedKelas === "Semua") {
      return recapData;
    }
    return recapData.filter(
      (item) => String(item.kelas).trim() === selectedKelas
    );
  }, [recapData, selectedKelas]);

  const getStatusSummary = (): {
    Hadir: number;
    Izin: number;
    Sakit: number;
    Alpha: number;
  } => {
    const summary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredRecapData.forEach((item) => {
      summary.Hadir += item.hadir || 0;
      summary.Alpha += item.alpa || 0;
      summary.Izin += item.izin || 0;
      summary.Sakit += item.sakit || 0;
    });
    return summary;
  };

  const statusSummary = getStatusSummary();

  const downloadExcel = () => {
    const headers = [
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const data = [
      headers,
      ...filteredRecapData.map((item) => [
        item.nama || "N/A",
        item.kelas || "N/A",
        item.hadir || 0,
        item.alpa || 0,
        item.izin || 0,
        item.sakit || 0,
        item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
      ]),
      [
        "TOTAL",
        "",
        statusSummary.Hadir,
        statusSummary.Alpha,
        statusSummary.Izin,
        statusSummary.Sakit,
        "",
      ],
      [
        "PERSEN",
        "",
        `${(
          (statusSummary.Hadir /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Alpha /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Izin /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Sakit /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        "",
      ],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = headers.map(() => ({ wch: 15 }));
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FFFF00" } },
      alignment: { horizontal: "center" },
    };
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    const percentStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    headers.forEach((header, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cellAddress] = { ...ws[cellAddress], s: headerStyle };
    });
    const totalRow = filteredRecapData.length + 1;
    ["A", "B", "C", "D", "E", "F", "G"].forEach((col, idx) => {
      const cellAddress = `${col}${totalRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: totalStyle };
    });
    const percentRow = filteredRecapData.length + 2;
    ["A", "B", "C", "D", "E", "F", "G"].forEach((col, idx) => {
      const cellAddress = `${col}${percentRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: percentStyle };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Semester");
    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Semester_${selectedSemester}_${selectedKelas}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const downloadPDF = () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    doc.setFont("Times", "roman");

    const title = `REKAP ABSENSI SISWA KELAS ${selectedKelas} SEMESTER ${selectedSemester} 2025`;
    doc.setFontSize(14);
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });
    currentY += 10;

    const headers = [
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const body = filteredRecapData.map((item) => [
      item.nama || "N/A",
      item.kelas || "N/A",
      item.hadir || 0,
      item.alpa || 0,
      item.izin || 0,
      item.sakit || 0,
      item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
    ]);

    const totalRow = [
      "TOTAL",
      "",
      statusSummary.Hadir,
      statusSummary.Alpha,
      statusSummary.Izin,
      statusSummary.Sakit,
      "",
    ];
    const percentRow = [
      "PERSEN",
      "",
      `${(
        (statusSummary.Hadir /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Alpha /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Izin /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Sakit /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      "",
    ];

    autoTable(doc, {
      head: [headers],
      body: [...body, totalRow, percentRow],
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      columnStyles: {
        0: { cellWidth: 50 },
        1: { cellWidth: 20 },
        2: { cellWidth: 20 },
        3: { cellWidth: 20 },
        4: { cellWidth: 20 },
        5: { cellWidth: 20 },
        6: { cellWidth: 20 },
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${placeName}, ${formattedDate}`;
      const rightColumnX = pageWidth - margin - 50;
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5;

      const principalText = [
        "Kepala Sekolah,",
        "",
        "",
        `( ${schoolData.namaKepsek || "N/A"} )`,
        `NIP: ${schoolData.nipKepsek || "N/A"}`,
      ];
      const teacherText = [
        "Guru Kelas,",
        "",
        "",
        `( ${schoolData.namaGuru || "N/A"} )`,
        `NIP: ${schoolData.nipGuru || "N/A"}`,
      ];

      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      // Principal signature and text
      if (schoolData.ttdKepsek) {
        doc.addImage(
          schoolData.ttdKepsek,
          "PNG",
          leftColumnX + 10,
          currentY - 3,
          signatureWidth,
          signatureHeight
        );
      }

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama kepala sekolah dengan format bold dan underline
      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to principal name
      const principalNameText = principalName;
      const textWidth = doc.getTextWidth(principalNameText);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );

      // Teacher signature and text
      if (schoolData.ttdGuru) {
        doc.addImage(
          schoolData.ttdGuru,
          "PNG",
          rightColumnX + 10,
          currentY - 5,
          signatureWidth,
          signatureHeight
        );
      }

      // Pisahkan "Guru Kelas" dengan posisi yang lebih tinggi
      doc.text("Guru Kelas,", rightColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama guru dengan format bold dan underline
      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to teacher name
      const teacherNameText = teacherName;
      const teacherTextWidth = doc.getTextWidth(teacherNameText);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Semester_${selectedSemester}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📊 Rekap Absensi Semester
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Pilih Semester</p>
            <select
              value={selectedSemester}
              onChange={(e) => setSelectedSemester(e.target.value as "1" | "2")}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              <option value="1">Semester 1</option>
              <option value="2">Semester 2</option>
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Rekap Semester pada
            File PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Nama Tempat</p>
              <input
                type="text"
                value={placeName}
                onChange={(e) => setPlaceName(e.target.value)}
                placeholder="Masukkan nama tempat"
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
            <div className="text-green-600 font-bold text-lg">
              {statusSummary.Hadir}
            </div>
            <div className="text-green-700 text-sm">Hadir</div>
          </div>
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
            <div className="text-yellow-600 font-bold text-lg">
              {statusSummary.Izin}
            </div>
            <div className="text-yellow-700 text-sm">Izin</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
            <div className="text-blue-600 font-bold text-lg">
              {statusSummary.Sakit}
            </div>
            <div className="text-blue-700 text-sm">Sakit</div>
          </div>
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
            <div className="text-red-600 font-bold text-lg">
              {statusSummary.Alpha}
            </div>
            <div className="text-red-700 text-sm">Alpha</div>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat rekap...</p>
          </div>
        ) : filteredRecapData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data rekap untuk Semester {selectedSemester} kelas{" "}
              {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Coba pilih kelas atau semester lain.
            </p>
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Nama
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Kelas
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Hadir
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Alpha
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Izin
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Sakit
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      % Hadir
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecapData.map((item, index) => (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.hadir || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.alpa || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.izin || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.sakit || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.persenHadir !== undefined
                          ? `${item.persenHadir}%`
                          : "N/A"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 flex gap-4 justify-center">
              <button
                onClick={downloadExcel}
                className="px-1 py-0.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download Excel
              </button>
              <button
                onClick={downloadPDF}
                className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📄 Download PDF
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const ClearDataTab: React.FC<{
  onRefresh?: () => void;
  onDataCleared?: () => void;
}> = ({ onRefresh, onDataCleared }) => {
  const [isClearing, setIsClearing] = useState<boolean>(false);

  const clearAllLocalData = () => {
    // Hapus semua data dari localStorage
    const keysToRemove = [
      "students",
      "studentData",
      "dataSiswa",
      "siswaData",
      "studentList",
      "daftarSiswa",
    ];
    keysToRemove.forEach((key) => {
      localStorage.removeItem(key);
    });

    // Hapus semua data dari sessionStorage
    sessionStorage.clear();

    // Hapus semua data yang mungkin ada dengan prefix tertentu
    for (let i = localStorage.length - 1; i >= 0; i--) {
      const key = localStorage.key(i);
      if (
        key &&
        (key.includes("student") ||
          key.includes("siswa") ||
          key.includes("data"))
      ) {
        localStorage.removeItem(key);
      }
    }

    console.log("Semua data lokal berhasil dihapus");
  };

  const handleClearData = async () => {
    if (
      !window.confirm(
        "Yakin ingin menghapus semua data di sheet Absensi dan DataSiswa?\n\nHeader akan tetap dipertahankan. Data siswa di aplikasi juga akan dihapus. Tindakan ini tidak dapat dibatalkan."
      )
    ) {
      return;
    }

    setIsClearing(true);

    // Langsung hapus data lokal dulu
    clearAllLocalData();

    // Buat AbortController untuk timeout
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000); // 30 detik timeout

    try {
      console.log("Mengirim request ke:", endpoint);
      console.log("Payload:", {
        type: "deleteAllDataDataSiswanAbsensi",
        sheet: "both",
      });

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          type: "deleteAllDataDataSiswanAbsensi",
          sheet: "both",
        }),
        // Gunakan AbortController untuk timeout
        signal: controller.signal,
      });

      // Clear timeout jika request berhasil
      clearTimeout(timeoutId);

      console.log("Status respons:", response.status);

      if (!response.ok) {
        throw new Error(
          `Server error: ${response.status} ${response.statusText}`
        );
      }

      const jsonResponse = await response.json();
      console.log("Respons JSON:", jsonResponse);

      if (jsonResponse.success) {
        // Pastikan data lokal benar-benar terhapus
        clearAllLocalData();

        // Trigger callback untuk parent component
        if (onDataCleared) {
          onDataCleared();
        }

        // Trigger event untuk memberitahu komponen lain bahwa data telah dihapus
        window.dispatchEvent(new CustomEvent("dataCleared"));

        // Trigger refresh untuk memuat ulang data dari server
        if (onRefresh) {
          onRefresh();
        }

        // Trigger event global untuk refresh
        window.dispatchEvent(new CustomEvent("refreshData"));

        alert(
          `✅ ${jsonResponse.message}\n\nData siswa di aplikasi juga telah dihapus. Halaman akan dimuat ulang.`
        );

        // Reload halaman untuk memastikan semua komponen direset
        setTimeout(() => {
          window.location.reload();
        }, 1000);
      } else {
        throw new Error(jsonResponse.message || "Gagal menghapus data");
      }
    } catch (error) {
      console.error("Error saat menghapus data:", error);

      const errorMessage =
        error instanceof Error
          ? error.message
          : "Terjadi kesalahan tidak diketahui";

      // Penanganan spesifik untuk error CORS atau jaringan
      if (
        errorMessage.includes("Failed to fetch") ||
        errorMessage.toLowerCase().includes("cors") ||
        errorMessage.includes("NetworkError")
      ) {
        console.warn(
          "Mendeteksi masalah CORS atau jaringan, mencoba fallback..."
        );
        try {
          await fetch(endpoint, {
            method: "POST",
            mode: "no-cors", // Fallback untuk CORS
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              type: "deleteAllDataDataSiswanAbsensi",
              sheet: "both",
            }),
          });
          console.log("Fallback request dikirim (no-cors)");

          // Pastikan data lokal benar-benar terhapus untuk fallback
          clearAllLocalData();

          // Trigger callback untuk parent component
          if (onDataCleared) {
            onDataCleared();
          }

          // Trigger event untuk memberitahu komponen lain bahwa data telah dihapus
          window.dispatchEvent(new CustomEvent("dataCleared"));

          // Trigger refresh untuk memuat ulang data dari server
          if (onRefresh) {
            onRefresh();
          }

          // Trigger event global untuk refresh
          window.dispatchEvent(new CustomEvent("refreshData"));

          alert(
            "✅ Data berhasil dihapus (CORS fallback). Data siswa di aplikasi juga telah dihapus. Halaman akan dimuat ulang."
          );

          // Reload halaman untuk memastikan semua komponen direset
          setTimeout(() => {
            window.location.reload();
          }, 1000);
        } catch (fallbackError) {
          console.error("Fallback error:", fallbackError);
          alert(
            `❌ Gagal menghapus data: ${
              fallbackError instanceof Error
                ? fallbackError.message
                : "Unknown error"
            }. Periksa koneksi jaringan atau endpoint.`
          );
        }
      } else if (
        errorMessage.includes("aborted") ||
        errorMessage.includes("timeout")
      ) {
        alert("❌ Gagal menghapus data: Permintaan timeout. Coba lagi nanti.");
      } else {
        alert(`❌ Gagal menghapus data: ${errorMessage}. Detail di console.`);
      }
    } finally {
      // Pastikan timeout dibersihkan
      clearTimeout(timeoutId);
      setIsClearing(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto pb-20">
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-red-700 mb-6">
          🗑️ Hapus Data
        </h2>
        <div className="bg-red-50 border border-red-200 rounded-lg p-4 mb-6">
          <p className="text-sm text-red-700 font-semibold mb-2">Peringatan:</p>
          <p className="text-sm text-red-600">
            Tindakan ini akan menghapus semua data di sheet Absensi dan
            DataSiswa (kecuali header), serta data siswa yang tersimpan di
            aplikasi. Tindakan ini tidak dapat dibatalkan.
          </p>
        </div>
        <div className="flex justify-center">
          <button
            onClick={handleClearData}
            disabled={isClearing}
            className={`px-6 py-2 rounded-lg font-medium text-white ${
              isClearing
                ? "bg-gray-400 cursor-not-allowed"
                : "bg-red-600 hover:bg-red-700"
            } transition-colors duration-200`}
          >
            {isClearing ? "Memproses..." : "🗑️ Hapus Semua Data"}
          </button>
        </div>
      </div>
    </div>
  );
};

// Komponen SplashScreen
const SplashScreen: React.FC = () => {
  return (
    <div className="fixed inset-0 bg-white flex flex-col items-center justify-center z-50">
      <style>
        {`
          @keyframes pulse {
            0% {
              transform: scale(1);
              opacity: 1;
            }
            50% {
              transform: scale(1.2);
              opacity: 0.7;
            }
            100% {
              transform: scale(1);
              opacity: 1;
            }
          }
          .animate-pulse-custom {
            animation: pulse 2s infinite;
          }
        `}
      </style>
      <img
        src="\images\IMG_20250518_064410.png"
        alt="Logo Aplikasi"
        className="w-52 h-70 mb-4 animate-pulse-custom" //Pengaturan ukuran logo
      />
      <p className="text-gray-800 text-lg font-semibold mt-6">
        Ronaldi, S.Pd
      </p>
    </div>
  );
};

const StudentAttendanceApp: React.FC = () => {
  const [students, setStudents] = useState<Student[]>([]);
  const [uniqueClasses, setUniqueClasses] = useState<string[]>(["Semua"]);
  const [activeTab, setActiveTab] = useState<
    | "schoolData"
    | "studentData"
    | "attendance"
    | "recap"
    | "graph"
    | "history"
    | "semesterRecap"
    | "clearData"
  >("studentData");
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  const fetchStudents = () => {
    fetch(endpoint)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data: Student[]) => {
        console.log("Data siswa yang diambil:", data);
        setStudents(data);

        const classSet = new Set<string>();
        data.forEach((student) => {
          if (student.kelas != null) {
            const kelasValue = String(student.kelas).trim();
            if (
              kelasValue !== "" &&
              kelasValue !== "undefined" &&
              kelasValue !== "null"
            ) {
              classSet.add(kelasValue);
            }
          }
        });
        const classes = Array.from(classSet).sort((a, b) => {
          const aIsNum = /^\d+$/.test(a);
          const bIsNum = /^\d+$/.test(b);
          if (aIsNum && bIsNum) return parseInt(a) - parseInt(b);
          if (aIsNum && !bIsNum) return -1;
          if (!aIsNum && bIsNum) return 1;
          return a.localeCompare(b);
        });
        setUniqueClasses(["Semua", ...classes]);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal mengambil data siswa. Cek console untuk detail.");
      });
  };

  const handleRecapRefresh = () => {
    setRefreshTrigger((prev) => prev + 1);
  };

  const handleRefresh = () => {
    fetchStudents();
  };

  useEffect(() => {
    // Simulasi loading selama 3 detik
    const timer = setTimeout(() => {
      setIsLoading(false);
      fetchStudents();
    }, 3000);

    return () => clearTimeout(timer);
  }, []);

  if (isLoading) {
    return <SplashScreen />;
  }

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col">
      {/* Sidebar */}
      <aside
        className={`bg-white shadow-md w-64 space-y-2 py-6 px-2 fixed h-full top-0 left-0 transform ${
          isSidebarOpen ? "translate-x-0" : "-translate-x-full"
        } transition-transform duration-300 ease-in-out z-50`}
      >
        <div className="flex justify-between items-center mb-4 px-4">
          <h2 className="text-xl font-bold text-gray-800">Menu</h2>
          <button
            onClick={() => setIsSidebarOpen(false)}
            className="text-gray-600 hover:text-gray-800 text-2xl"
          >
            ✖️
          </button>
        </div>
        {[
          { tab: "schoolData", label: "🏫 Data Sekolah" },
          { tab: "studentData", label: "👥 Data Siswa" },
          { tab: "attendance", label: "📋 Absensi" },
          { tab: "recap", label: "📊 Rekap Bulanan" },
          { tab: "semesterRecap", label: "📚 Rekap Semester" },
          { tab: "graph", label: "📈 Grafik" },
          { tab: "history", label: "📜 Riwayat Absen" },
          { tab: "clearData", label: "🗑️ Hapus Data" },
        ].map(({ tab, label }) => (
          <button
            key={tab}
            onClick={() => {
              setActiveTab(
                tab as
                  | "schoolData"
                  | "studentData"
                  | "attendance"
                  | "recap"
                  | "graph"
                  | "history"
                  | "semesterRecap"
                  | "clearData"
              );
              setIsSidebarOpen(false);
            }}
            className={`w-full text-left py-2 px-4 rounded-lg text-sm font-medium transition-colors flex items-center gap-2 ${
              activeTab === tab
                ? "bg-blue-600 text-white"
                : "text-gray-600 hover:bg-gray-100"
            }`}
          >
            {label}
          </button>
        ))}
      </aside>

      {/* Hamburger Menu Button */}
      <div className="fixed top-4 left-4 z-50">
        <button
          onClick={() => setIsSidebarOpen(!isSidebarOpen)}
          className="bg-blue-600 text-white px-4 py-2 rounded-lg shadow-md hover:bg-blue-700 transition-colors"
        >
          {isSidebarOpen ? "✖️ Tutup Menu" : "☰ Buka Menu"}
        </button>
      </div>

      {/* Logo di pojok kanan atas */}
      <div className="absolute top-4 right-4 z-50">
        <img
          src="\images\Untitled design (6).png"
          alt="Logo Aplikasi"
          className="w-16 h-16"
        />
      </div>

      {/* Main Content */}
      <main
        className={`flex-1 p-6 transition-all duration-300 ${
          isSidebarOpen ? "ml-64" : "ml-0"
        } mt-16`}
      >
        <div className="text-center mb-6">
          <h1 className="text-3xl font-bold text-gray-800 mb-2">
            Sistem Absensi Siswa
          </h1>
          <p className="text-gray-600">Kelola data siswa dan absensi harian</p>
        </div>

        <div className="py-4">
          {activeTab === "schoolData" && (
            <SchoolDataTab onRefresh={handleRefresh} />
          )}
          {activeTab === "studentData" && (
            <StudentDataTab
              students={students}
              onRefresh={fetchStudents}
              uniqueClasses={uniqueClasses}
            />
          )}
          {activeTab === "attendance" && (
            <AttendanceTab
              students={students}
              onRecapRefresh={handleRecapRefresh}
            />
          )}
          {activeTab === "recap" && (
            <MonthlyRecapTab
              onRefresh={handleRecapRefresh}
              uniqueClasses={uniqueClasses}
            />
          )}
          {activeTab === "graph" && <GraphTab uniqueClasses={uniqueClasses} />}
          {activeTab === "history" && (
            <AttendanceHistoryTab
              students={students}
              uniqueClasses={uniqueClasses}
              onRefresh={fetchStudents}
            />
          )}
          {activeTab === "semesterRecap" && (
            <SemesterRecapTab uniqueClasses={uniqueClasses} />
          )}
          {activeTab === "clearData" && <ClearDataTab />}
        </div>
      </main>
    </div>
  );
};

export default StudentAttendanceApp;
