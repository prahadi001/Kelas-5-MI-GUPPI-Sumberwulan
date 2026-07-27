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

// ===== Postingan terbaru di Beranda =====
async function muatPostinganTerbaru() {
  const wadah = document.getElementById('posting-terbaru');
  if (!wadah) return;

  try {
    const res = await fetch('posts/beranda-posts.json');
    if (!res.ok) throw new Error('Gagal memuat data');
    const daftar = await res.json();

    const lima = daftar.slice(0, 5);
    if (lima.length === 0) {
      wadah.innerHTML = '<p class="status-muat">Belum ada postingan.</p>';
      return;
    }

    wadah.innerHTML = lima.map(p => `
      <a class="posting-kartu" href="#">
        <div class="posting-thumb">${p.ikon || '📰'}</div>
        <div class="posting-isi">
          <div class="posting-kategori">${p.kategori || 'Postingan'}</div>
          <h3>${p.judul}</h3>
          <p class="tanggal">${p.tanggal}</p>
          <p class="ringkasan">${p.ringkasan}</p>
        </div>
      </a>
    `).join('');

  } catch (err) {
    wadah.innerHTML = '<p class="status-muat">Postingan belum bisa dimuat. Jika halaman ini baru saja diunggah ke GitHub Pages, coba muat ulang beberapa saat lagi.</p>';
  }
}
document.addEventListener('DOMContentLoaded', muatPostinganTerbaru);

// ===== Slider foto di Beranda =====
function jalankanSlider() {
  const slides = document.querySelectorAll('.slide');
  const titikWadah = document.querySelector('.slider-titik');
  if (slides.length === 0) return;

  let index = 0;

  function tampilkan(i) {
    slides.forEach(s => s.classList.remove('aktif'));
    document.querySelectorAll('.slider-titik button').forEach(b => b.classList.remove('aktif'));
    slides[i].classList.add('aktif');
    if (titikWadah) titikWadah.children[i].classList.add('aktif');
    index = i;
  }

  if (titikWadah) {
    slides.forEach((_, i) => {
      const btn = document.createElement('button');
      if (i === 0) btn.classList.add('aktif');
      btn.setAttribute('aria-label', 'Slide ' + (i + 1));
      btn.addEventListener('click', () => tampilkan(i));
      titikWadah.appendChild(btn);
    });
  }

  document.querySelector('.slider-panah.kanan')?.addEventListener('click', () => {
    tampilkan((index + 1) % slides.length);
  });
  document.querySelector('.slider-panah.kiri')?.addEventListener('click', () => {
    tampilkan((index - 1 + slides.length) % slides.length);
  });

  setInterval(() => {
    tampilkan((index + 1) % slides.length);
  }, 4500);
}
document.addEventListener('DOMContentLoaded', jalankanSlider);
