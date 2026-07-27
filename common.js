// Mengambil daftar tugas secara otomatis dari posts/posts.json
// dan menampilkannya di dalam elemen dengan id="daftar-tugas"

async function muatDaftarTugas() {
  const wadah = document.getElementById('daftar-tugas');
  if (!wadah) return;

  try {
    const res = await fetch('posts/posts.json');
    if (!res.ok) throw new Error('Gagal memuat data');
    const daftar = await res.json();

    if (!Array.isArray(daftar) || daftar.length === 0) {
      wadah.innerHTML = '<p class="status-muat">Belum ada tugas yang diberikan.</p>';
      return;
    }

    wadah.innerHTML = daftar.map(tugas => `
      <div class="item">
        <span class="label-pill">Tugas</span>
        <h3><a href="#">${tugas.judul}</a></h3>
        <p class="tenggat">Tenggat: ${tugas.tenggat}</p>
        <p class="deskripsi">${tugas.deskripsi}</p>
      </div>
    `).join('');

  } catch (err) {
    wadah.innerHTML = '<p class="status-muat">Daftar tugas belum bisa dimuat. Jika halaman ini baru saja diunggah ke GitHub Pages, coba muat ulang beberapa saat lagi.</p>';
  }
}

document.addEventListener('DOMContentLoaded', muatDaftarTugas);
