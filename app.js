
/* Rirekisho CV Generator
   Form, Live Preview, and Excel export synchronization.
*/
(() => {
  'use strict';

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));

  const form = $('#cvForm');
  const toast = $('#toast');
  const downloadButton = $('#downloadExcel');
  const downloadStatus = $('#downloadStatus');

  const limits = {
    education: 4,
    work: 4,
    certificate: 3,
    family: 6
  };

  const containers = {
    education: $('#educationRows'),
    work: $('#workRows'),
    certificate: $('#certificateRows'),
    family: $('#familyRows')
  };

  let photoDataUrl = '';
  let furiganaWasEdited = false;
  let toastTimer = null;

  // ==========================================
  // 1. KONVERSI ROMAJI KE KATAKANA
  // ==========================================

  const KANA = {
    kya:'キャ', kyu:'キュ', kyo:'キョ',
    gya:'ギャ', gyu:'ギュ', gyo:'ギョ',
    sha:'シャ', shu:'シュ', sho:'ショ',
    cha:'チャ', chu:'チュ', cho:'チョ',
    ja:'ジャ', ju:'ジュ', jo:'ジョ',
    nya:'ニャ', nyu:'ニュ', nyo:'ニョ',
    hya:'ヒャ', hyu:'ヒュ', hyo:'ヒョ',
    bya:'ビャ', byu:'ビュ', byo:'ビョ',
    pya:'ピャ', pyu:'ピュ', pyo:'ピョ',
    rya:'リャ', ryu:'リュ', ryo:'リョ',
    mya:'ミャ', myu:'ミュ', myo:'ミョ',

    shi:'シ', chi:'チ', tsu:'ツ', fu:'フ', ji:'ジ',
    fa:'ファ', fi:'フィ', fe:'フェ', fo:'フォ',
    va:'ヴァ', vi:'ヴィ', vu:'ヴ', ve:'ヴェ', vo:'ヴォ',
    wi:'ウィ', we:'ウェ', wo:'ヲ',
    she:'シェ', je:'ジェ', che:'チェ', tse:'ツェ',

    a:'ア', i:'イ', u:'ウ', e:'エ', o:'オ',
    ka:'カ', ki:'キ', ku:'ク', ke:'ケ', ko:'コ',
    ga:'ガ', gi:'ギ', gu:'グ', ge:'ゲ', go:'ゴ',
    sa:'サ', si:'シ', su:'ス', se:'セ', so:'ソ',
    za:'ザ', zi:'ジ', zu:'ズ', ze:'ゼ', zo:'ゾ',
    ta:'タ', ti:'ティ', tu:'トゥ', te:'テ', to:'ト',
    da:'ダ', di:'ディ', du:'ドゥ', de:'デ', do:'ド',
    na:'ナ', ni:'ニ', nu:'ヌ', ne:'ネ', no:'ノ',
    ha:'ハ', hi:'ヒ', hu:'フ', he:'ヘ', ho:'ホ',
    ba:'バ', bi:'ビ', bu:'ブ', be:'ベ', bo:'ボ',
    pa:'パ', pi:'ピ', pu:'プ', pe:'ペ', po:'ポ',
    ma:'マ', mi:'ミ', mu:'ム', me:'メ', mo:'モ',
    ya:'ヤ', yu:'ユ', yo:'ヨ',
    ra:'ラ', ri:'リ', ru:'ル', re:'レ', ro:'ロ',
    wa:'ワ',
    la:'ラ', li:'リ', lu:'ル', le:'レ', lo:'ロ',
    xa:'ァ', xi:'ィ', xu:'ゥ', xe:'ェ', xo:'ォ'
  };

  function romajiToKatakana(raw) {
    if (!raw || !raw.trim()) return '';

    const normalized = raw
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/kh/g, 'k')
      .replace(/shy/g, 'sh')
      .replace(/sy/g, 'sh')
      .replace(/dh/g, 'd')
      .replace(/th/g, 't')
      .replace(/dz/g, 'z')
      .replace(/ph/g, 'f')
      .replace(/c(?=[eiy])/g, 'ch')
      .replace(/c/g, 'k')
      .replace(/q/g, 'k')
      .replace(/[^a-z\s.'-]/g, ' ');

    const tokens = normalized.split(/[\s.'-]+/).filter(Boolean);

    const fallback = {
      b:'ブ', c:'ク', d:'ド', f:'フ', g:'グ',
      h:'フ', j:'ジ', k:'ク', l:'ル', m:'ム',
      p:'プ', q:'ク', r:'ル', s:'ス', t:'ト',
      v:'ヴ', w:'ウ', x:'クス', z:'ズ'
    };

    return tokens.map(token => {
      let output = '';
      let i = 0;

      while (i < token.length) {
        const current = token[i];
        const next = token[i + 1] || '';

        // Konsonan ganda
        if (current === next && !'aeioun'.includes(current)) {
          output += 'ッ';
          i++;
          continue;
        }

        // Huruf N
        if (
          current === 'n' &&
          (
            i === token.length - 1 ||
            (next && !'aiueoy'.includes(next))
          )
        ) {
          output += 'ン';
          i++;
          continue;
        }

        let match = '';

        for (let size = Math.min(3, token.length - i); size >= 1; size--) {
          const candidate = token.slice(i, i + size);

          if (Object.prototype.hasOwnProperty.call(KANA, candidate)) {
            match = candidate;
            break;
          }
        }

        if (match) {
          output += KANA[match];
          i += match.length;
          continue;
        }

        if (/[a-z]/.test(current)) {
          if (current === 'n') {
            output += 'ン';
          } else if (fallback[current]) {
            output += fallback[current];
          }

          i++;
        } else {
          i++;
        }
      }

      return output;
    }).join('・');
  }

  // ==========================================
  // 2. FUNGSI DASAR
  // ==========================================

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, char => ({
      '&':'&amp;',
      '<':'&lt;',
      '>':'&gt;',
      '"':'&quot;',
      "'":'&#39;'
    })[char]);
  }

  function value(id) {
    return $(`#${id}`)?.value?.trim() || '';
  }

  function genderToJapanese(gender) {
    const normalized = String(gender || '').trim().toLowerCase();

    if (normalized === '男性' || normalized === 'laki-laki') {
      return '男性';
    }

    if (normalized === '女性' || normalized === 'perempuan') {
      return '女性';
    }

    return '';
  }

  function calculateAge(dateValue) {
    if (!dateValue) return '';

    const parts = dateValue.split('-').map(Number);

    if (parts.length !== 3 || parts.some(number => !number)) {
      return '';
    }

    const [year, month, day] = parts;
    const today = new Date();

    let age = today.getFullYear() - year;

    if (
      today.getMonth() + 1 < month ||
      (
        today.getMonth() + 1 === month &&
        today.getDate() < day
      )
    ) {
      age--;
    }

    return age >= 0 ? age : '';
  }

  function dateParts(dateValue) {
    if (!dateValue) return null;

    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateValue);

    if (!match) return null;

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const date = new Date(year, month - 1, day);

    if (
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day
    ) {
      return null;
    }

    return { year, month, day };
  }

  function formatBirthDate(dateValue) {
    const p = dateParts(dateValue);

    return p
      ? `${p.year}年${p.month}月${p.day}日`
      : '';
  }

  function formatMonthYear(year, month) {
    if (!year && !month) return '';

    return [
      year ? `${year}年` : '',
      month ? `${Number(month)}月` : ''
    ].filter(Boolean).join(' ');
  }

  function formatCertDate(dateValue) {
    const p = dateParts(dateValue);

    return p
      ? `${p.year}年 ${p.month}月${p.day}日`
      : '';
  }

  function formatCreatedDate() {
    const now = new Date();

    return `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;
  }

  function cleanFilename(name) {
    return (name || 'Pengguna')
      .normalize('NFKD')
      .replace(/[^a-zA-Z0-9_-]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .slice(0, 48) || 'Pengguna';
  }

  function showToast(message) {
    if (!toast) return;

    toast.textContent = message;
    toast.classList.add('show');

    window.clearTimeout(toastTimer);

    toastTimer = window.setTimeout(() => {
      toast.classList.remove('show');
    }, 3200);
  }

  function updateAgeDisplay() {
    const age = calculateAge(value('birthDate'));
    $('#ageValue').textContent = age === '' ? '—' : age;
  }

  // ==========================================
  // 3. DATA FORM BERULANG
  // ==========================================

  const rowDefinitions = {
    education: {
      title:'Pendidikan',
      fields:[
        ['startYear','Tahun masuk','number','min="1900" max="2200" placeholder="2019"'],
        ['startMonth','Bulan masuk','number','min="1" max="12" placeholder="7"'],
        ['endYear','Tahun lulus','number','min="1900" max="2200" placeholder="2023"'],
        ['endMonth','Bulan lulus','number','min="1" max="12" placeholder="6"'],
        ['schoolName','Nama sekolah / universitas','text','placeholder="Nama institusi pendidikan"','full'],
        ['schoolType','Jenjang pendidikan','select','options']
      ]
    },

    work: {
      title:'Pekerjaan',
      fields:[
        ['startYear','Tahun masuk','number','min="1900" max="2200" placeholder="2023"'],
        ['startMonth','Bulan masuk','number','min="1" max="12" placeholder="1"'],
        ['endYear','Tahun keluar','number','min="1900" max="2200" placeholder="2025"'],
        ['endMonth','Bulan keluar','number','min="1" max="12" placeholder="12"'],
        ['companyName','Nama perusahaan / tempat kerja','text','placeholder="Nama perusahaan"','full'],
        ['residenceStatus','Status izin tinggal','text','placeholder="Contoh: Indonesia / Tokutei Ginou"'],
        ['workDescription','Deskripsi pekerjaan','text','placeholder="Tugas dan tanggung jawab"']
      ]
    },

    certificate: {
      title:'Sertifikat',
      fields:[
        ['date','Tanggal perolehan','date',''],
        ['certificateName','Nama sertifikat / ujian','text','placeholder="Contoh: JFT-Basic A2"'],
        ['japaneseName','Nama sertifikat dalam bahasa Jepang','text','placeholder="Contoh: 国際交流基金日本語基礎テスト"','full']
      ]
    },

    family: {
      title:'Anggota keluarga',
      fields:[
        ['relation','Hubungan keluarga','text','placeholder="Ayah / Ibu / Kakak"'],
        ['name','Nama lengkap','text','placeholder="Nama anggota keluarga"'],
        ['age','Usia','number','min="0" max="120" placeholder="45"'],
        ['address','Alamat','text','placeholder="Alamat anggota keluarga"','full'],
        ['occupation','Pekerjaan','text','placeholder="Pekerjaan anggota keluarga"','full']
      ]
    }
  };

  function fieldMarkup(type, fieldDef) {
    const [name, label, inputType, extra = '', width = ''] = fieldDef;

    const full = width === 'full' ? ' full' : '';

    if (inputType === 'select' && name === 'schoolType') {
      return `
        <div class="repeat-field${full}">
          <label>Jenjang pendidikan</label>
          <select data-row-field="schoolType">
            <option value="小学校">SD / 小学校</option>
            <option value="中学校">SMP / 中学校</option>
            <option value="高等学校">SMA/SMK / 高等学校</option>
            <option value="大学">Universitas / 大学</option>
            <option value="専門学校">Sekolah kejuruan / 専門学校</option>
            <option value="その他">Lainnya / その他</option>
          </select>
        </div>
      `;
    }

    return `
      <div class="repeat-field${full}">
        <label>${label}</label>
        <input
          type="${inputType}"
          data-row-field="${name}"
          ${extra}
        />
      </div>
    `;
  }

  function addRow(type) {
    const container = containers[type];

    if (!container || !rowDefinitions[type]) return;

    const current = $$('.repeat-item', container).length;

    if (current >= limits[type]) {
      showToast(`Maksimal ${limits[type]} data untuk bagian ini.`);
      return;
    }

    const definition = rowDefinitions[type];

    const item = document.createElement('div');

    item.className = 'repeat-item';
    item.dataset.rowType = type;

    const content = definition.fields
      .map(fieldDef => fieldMarkup(type, fieldDef))
      .join('');

    item.innerHTML = `
      <div class="repeat-head">
        <strong>${definition.title} ${current + 1}</strong>
        <button
          type="button"
          class="remove-row"
          aria-label="Hapus ${definition.title}"
          title="Hapus"
        >×</button>
      </div>

      <div class="repeat-grid">
        ${content}
      </div>
    `;

    container.append(item);

    updateAddButton(type);
    renderPreview();
  }

  function updateAddButton(type) {
    const btn = $(`[data-add="${type}"]`);
    const container = containers[type];

    if (!btn || !container) return;

    const count = $$('.repeat-item', container).length;

    const names = {
      education:'pendidikan',
      work:'pekerjaan',
      certificate:'sertifikat',
      family:'anggota keluarga'
    };

    btn.disabled = count >= limits[type];

    btn.textContent = count >= limits[type]
      ? `Maksimal ${limits[type]} data tercapai`
      : `＋ Tambah ${names[type]}`;

    $$('.repeat-item', container).forEach((item, index) => {
      const label = $('.repeat-head strong', item);

      if (label) {
        label.textContent = `${rowDefinitions[type].title} ${index + 1}`;
      }
    });
  }

  function readRows(type) {
    const container = containers[type];

    if (!container) return [];

    return $$('.repeat-item', container)
      .map(item => {
        const row = {};

        $$('[data-row-field]', item).forEach(input => {
          row[input.dataset.rowField] = input.value.trim();
        });

        return row;
      })
      .filter(row => Object.values(row).some(Boolean));
  }

  // ==========================================
  // 4. MENGAMBIL SELURUH DATA
  // ==========================================

  function collectData() {
    const visited = value('visitedJapan') === 'yes';

    return {
      fullName: value('fullName'),
      furigana: value('furigana'),

      birthDate: value('birthDate'),
      age: calculateAge(value('birthDate')),

      gender: genderToJapanese(value('gender')),
      maritalStatus: value('maritalStatus') || '無',

      birthPlace: value('birthPlace'),
      heightCm: value('heightCm'),
      weightKg: value('weightKg'),

      // Menggunakan BULAN, bukan hari.
      tokuteiMonths: value('tokuteiMonths'),

      address: value('address'),
      hobby: value('hobby'),

      visitedJapan: visited,
      travelPeriod: visited ? value('travelPeriod') : '無し',
      travelPurpose: visited ? value('travelPurpose') : '無し',

      education: readRows('education'),
      work: readRows('work'),
      certificates: readRows('certificate'),
      family: readRows('family'),

      motivation: value('motivation'),
      photo: photoDataUrl
    };
  }

  // ==========================================
  // 5. MEMBUAT LIVE PREVIEW
  // ==========================================

  function monthCell(year, month) {
    return escapeHtml(formatMonthYear(year, month));
  }

  function sectionLabel(japanese, indonesian) {
    return `
      <div class="cv-section">
        ${escapeHtml(japanese)}
        <br>
        <span>${escapeHtml(indonesian)}</span>
      </div>
    `;
  }

  function displayOrDash(text) {
    return text
      ? escapeHtml(text)
      : '<span class="empty-value">—</span>';
  }

  function renderPreview() {
    const preview = $('#cvPreview');

    if (!preview) return;

    const data = collectData();

    updateAgeDisplay();

    const photo = data.photo
      ? `<img src="${data.photo}" alt="Foto profil" />`
      : `
        <div class="photo-placeholder-small">
          写真
          <br>
          <small>3 × 4 cm</small>
        </div>
      `;

    const educationRows = Array.from({ length:4 }, (_, index) => {
      const row = data.education[index] || {};

      return `
        <tr>
          <td class="center">${monthCell(row.startYear, row.startMonth)}</td>
          <td class="center">${monthCell(row.endYear, row.endMonth)}</td>
          <td>${displayOrDash(row.schoolName)}</td>
          <td class="center">${displayOrDash(row.schoolType)}</td>
        </tr>
      `;
    }).join('');

    const workRows = Array.from({ length:4 }, (_, index) => {
      const row = data.work[index] || {};

      return `
        <tr>
          <td class="center">${monthCell(row.startYear, row.startMonth)}</td>
          <td class="center">${monthCell(row.endYear, row.endMonth)}</td>
          <td>${displayOrDash(row.companyName)}</td>
          <td class="center">${displayOrDash(row.residenceStatus)}</td>
          <td>${displayOrDash(row.workDescription)}</td>
        </tr>
      `;
    }).join('');

    const certificateRows = Array.from({ length:3 }, (_, index) => {
      const row = data.certificates[index] || {};
      const date = formatCertDate(row.date);

      return `
        <tr>
          <td class="center">${displayOrDash(date)}</td>
          <td>${displayOrDash(row.certificateName)}</td>
          <td class="center">${displayOrDash(date)}</td>
          <td>${displayOrDash(row.japaneseName)}</td>
        </tr>
      `;
    }).join('');

    const familyRows = Array.from({ length:6 }, (_, index) => {
      const row = data.family[index] || {};

      return `
        <tr>
          <td class="center">${row.name || row.relation ? index + 1 : '—'}</td>
          <td>${displayOrDash(row.relation)}</td>
          <td>${displayOrDash(row.name)}</td>
          <td class="center">${row.age ? `${escapeHtml(row.age)}歳` : '—'}</td>
          <td>${displayOrDash(row.address)}</td>
          <td>${displayOrDash(row.occupation)}</td>
        </tr>
      `;
    }).join('');

    preview.innerHTML = `
      <div class="cv-top">
        <div class="cv-title">履歴書</div>
        <div class="cv-number">
          <div>NO</div>
          <div>作成日　${formatCreatedDate()}</div>
        </div>
      </div>

      <table class="cv-table personal-table">
        <colgroup>
          <col style="width:13%">
          <col style="width:13%">
          <col style="width:8%">
          <col style="width:7%">
          <col style="width:10%">
          <col style="width:8%">
          <col style="width:10%">
          <col style="width:8%">
          <col style="width:23%">
        </colgroup>

        <tr>
          <th class="cv-label">氏名<br>(Nama)</th>
          <td colspan="7">${displayOrDash(data.fullName)}</td>
          <td class="cv-photo" rowspan="4">${photo}</td>
        </tr>

        <tr>
          <th class="cv-label">フリガナ<br>(Furigana)</th>
          <td colspan="7">${displayOrDash(data.furigana)}</td>
        </tr>

        <tr>
          <th class="cv-label">生年月日<br>(Tanggal Lahir)</th>
          <td>${displayOrDash(formatBirthDate(data.birthDate))}</td>
          <th class="cv-sub-label">年齢<br>(Usia)</th>
          <td class="center">${data.age || '—'}</td>
          <th class="cv-sub-label">性別<br>(Jenis Kelamin)</th>
          <td class="center">${displayOrDash(data.gender)}</td>
          <th class="cv-sub-label">配偶者<br>(Suami/Istri)</th>
          <td class="center">${data.maritalStatus === '有' ? '有' : '無'}</td>
        </tr>

        <tr>
          <th class="cv-label">出身地<br>(Tempat Lahir)</th>
          <td>${displayOrDash(data.birthPlace)}</td>
          <th class="cv-sub-label">身長<br>(Tinggi)</th>
          <td class="center">${data.heightCm ? `${escapeHtml(data.heightCm)}cm` : '—'}</td>
          <th class="cv-sub-label">体重<br>(Berat)</th>
          <td class="center">${data.weightKg ? `${escapeHtml(data.weightKg)}kg` : '—'}</td>
          <th class="cv-sub-label">特定技能残り期間<br>(Bulan)</th>
          <td class="center">${data.tokuteiMonths ? `${escapeHtml(data.tokuteiMonths)}ヶ月` : '—'}</td>
        </tr>

        <tr>
          <th class="cv-label">現住所<br>(Alamat)</th>
          <td colspan="8">${displayOrDash(data.address)}</td>
        </tr>

        <tr>
          <th class="cv-label">趣味・特技<br>(Hobby)</th>
          <td colspan="8">${displayOrDash(data.hobby)}</td>
        </tr>
      </table>

      ${sectionLabel('日本への渡航履歴','Pernah ke Jepang?')}

      <table class="cv-table">
        <tr>
          <th class="cv-label">渡航時期<br>(Kapan)</th>
          <td style="width:40%">${displayOrDash(data.travelPeriod)}</td>
          <th class="cv-sub-label" style="width:20%">渡航目的<br>(Tujuan)</th>
          <td>${displayOrDash(data.travelPurpose)}</td>
        </tr>
      </table>

      ${sectionLabel('学歴','Riwayat Pendidikan')}

      <table class="cv-table">
        <tr class="cv-header">
          <th style="width:17%">入学年月<br>(Tahun Masuk)</th>
          <th style="width:17%">卒業年月<br>(Tahun Keluar)</th>
          <th>学校名<br>(Riwayat Pendidikan)</th>
          <th style="width:18%">学歴<br>(Jenis Sekolah)</th>
        </tr>
        ${educationRows}
      </table>

      ${sectionLabel('職歴','Riwayat Pekerjaan')}

      <table class="cv-table">
        <tr class="cv-header">
          <th style="width:14%">入社年月<br>(Tahun Masuk)</th>
          <th style="width:14%">退職年月<br>(Tahun Keluar)</th>
          <th style="width:29%">会社名<br>(Riwayat Pekerjaan)</th>
          <th style="width:19%">在留資格<br>(Status Izin Tinggal)</th>
          <th>業務内容<br>(Deskripsi Pekerjaan)</th>
        </tr>
        ${workRows}
      </table>

      ${sectionLabel('資格・免許','Sertifikat Keahlian')}

      <table class="cv-table">
        <tr class="cv-header">
          <th style="width:18%">取得年月<br>(Tanggal)</th>
          <th style="width:32%">資格<br>(Nama Ujian)</th>
          <th style="width:18%">取得年月</th>
          <th>資格<br>(Kualifikasi)</th>
        </tr>
        ${certificateRows}
      </table>

      ${sectionLabel('家族構成','Keluarga')}

      <table class="cv-table">
        <tr class="cv-header">
          <th style="width:5%">No</th>
          <th style="width:16%">続柄<br>(Hubungan)</th>
          <th style="width:20%">氏名<br>(Nama)</th>
          <th style="width:9%">年齢<br>(Usia)</th>
          <th style="width:32%">住所<br>(Alamat)</th>
          <th>職業<br>(Pekerjaan)</th>
        </tr>
        ${familyRows}
      </table>

      ${sectionLabel('志望動機','Alasan Melamar Pekerjaan')}

      <div class="cv-motivation">
        ${data.motivation
          ? escapeHtml(data.motivation).replace(/\n/g, '<br>')
          : '<span class="empty-value">Tuliskan motivasi melamar pekerjaan di sini.</span>'
        }
      </div>
    `;
  }

  // ==========================================
  // 6. FOTO 3 X 4
  // ==========================================

  async function cropPhoto(file) {
    if (!file || !/^image\/(jpeg|png)$/.test(file.type)) {
      throw new Error('Gunakan foto JPG atau PNG.');
    }

    const imageUrl = URL.createObjectURL(file);

    try {
      const img = new Image();

      img.src = imageUrl;

      await img.decode();

      const targetRatio = 3 / 4;

      let sx = 0;
      let sy = 0;
      let sw = img.width;
      let sh = img.height;

      if (sw / sh > targetRatio) {
        sw = sh * targetRatio;
        sx = (img.width - sw) / 2;
      } else {
        sh = sw / targetRatio;
        sy = (img.height - sh) / 2;
      }

      const canvas = document.createElement('canvas');

      canvas.width = 600;
      canvas.height = 800;

      const context = canvas.getContext('2d');

      if (!context) {
        throw new Error('Browser tidak mendukung pemrosesan foto.');
      }

      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, 600, 800);

      context.drawImage(
        img,
        sx, sy, sw, sh,
        0, 0, 600, 800
      );

      return canvas.toDataURL('image/jpeg', 0.92);

    } finally {
      URL.revokeObjectURL(imageUrl);
    }
  }

  // ==========================================
  // 7. MEMBUAT FILE EXCEL
  // ==========================================

  function setCell(ws, address, cellValue) {
    ws.getCell(address).value = cellValue ?? '';
  }

  function makeExcelWorkbook(data) {
    const workbook = new ExcelJS.Workbook();

    return workbook.xlsx.load(data.templateBuffer).then(() => {

      workbook.creator = 'Rirekisho CV Generator';
      workbook.subject = '履歴書';
      workbook.title = `履歴書 - ${data.fullName || 'CV'}`;
      workbook.modified = new Date();

      // Hapus sheet selain CV.
      const extraSheets = workbook.worksheets.filter(
        sheet => sheet.name !== 'CV'
      );

      extraSheets.forEach(sheet => {
        workbook.removeWorksheet(sheet.id);
      });

      const ws = workbook.getWorksheet('CV');

      if (!ws) {
        throw new Error('Sheet CV tidak ditemukan pada template.');
      }

      // IDENTITAS PRIBADI
      setCell(ws, 'M1', '');
      setCell(ws, 'M2', formatCreatedDate());

      setCell(ws, 'B3', data.fullName);
      setCell(ws, 'B4', data.furigana);
      setCell(ws, 'B5', formatBirthDate(data.birthDate));
      setCell(ws, 'E5', data.age);
      setCell(ws, 'H5', data.gender);
      setCell(ws, 'K5', data.maritalStatus);

      setCell(ws, 'B6', data.birthPlace);
      setCell(ws, 'E6', data.heightCm);
      setCell(ws, 'H6', data.weightKg);

      // TOKUTEI GINOU: BULAN
      setCell(
        ws,
        'K6',
        data.tokuteiMonths
          ? `${data.tokuteiMonths}ヶ月`
          : ''
      );

      setCell(ws, 'B8', data.address);
      setCell(ws, 'B10', data.hobby);

      // RIWAYAT PERJALANAN JEPANG
      setCell(ws, 'B12', data.travelPeriod);
      setCell(ws, 'I12', data.travelPurpose);

      // PENDIDIKAN
      for (let i = 0; i < 4; i++) {
        const rowNumber = 15 + i;
        const row = data.education[i] || {};

        setCell(
          ws,
          `A${rowNumber}`,
          formatMonthYear(row.startYear, row.startMonth)
        );

        setCell(
          ws,
          `B${rowNumber}`,
          formatMonthYear(row.endYear, row.endMonth)
        );

        setCell(ws, `C${rowNumber}`, row.schoolName);
        setCell(ws, `K${rowNumber}`, row.schoolType);
      }

      // PEKERJAAN
      for (let i = 0; i < 4; i++) {
        const rowNumber = 20 + i;
        const row = data.work[i] || {};

        setCell(
          ws,
          `A${rowNumber}`,
          formatMonthYear(row.startYear, row.startMonth)
        );

        setCell(
          ws,
          `B${rowNumber}`,
          formatMonthYear(row.endYear, row.endMonth)
        );

        setCell(ws, `C${rowNumber}`, row.companyName);
        setCell(ws, `K${rowNumber}`, row.residenceStatus);
        setCell(ws, `M${rowNumber}`, row.workDescription);
      }

      // SERTIFIKAT
      for (let i = 0; i < 3; i++) {
        const rowNumber = 26 + i;
        const row = data.certificates[i] || {};

        const acquiredDate = formatCertDate(row.date);

        setCell(ws, `A${rowNumber}`, acquiredDate);
        setCell(ws, `C${rowNumber}`, row.certificateName);
        setCell(ws, `G${rowNumber}`, acquiredDate);
        setCell(ws, `J${rowNumber}`, row.japaneseName);
      }

      // KELUARGA
      for (let i = 0; i < 6; i++) {
        const rowNumber = 31 + i;
        const row = data.family[i] || {};

        setCell(
          ws,
          `A${rowNumber}`,
          row.name || row.relation ? i + 1 : ''
        );

        setCell(ws, `B${rowNumber}`, row.relation);
        setCell(ws, `C${rowNumber}`, row.name);

        setCell(
          ws,
          `G${rowNumber}`,
          row.age ? `${row.age}歳` : ''
        );

        setCell(ws, `H${rowNumber}`, row.address);
        setCell(ws, `M${rowNumber}`, row.occupation);
      }

      // MOTIVASI
      setCell(ws, 'A38', data.motivation);

      // FOTO
      if (data.photo) {
        const imageId = workbook.addImage({
          base64: data.photo,
          extension: 'jpeg'
        });

        ws.addImage(imageId, {
          tl: { col:12.06, row:2.05 },
          br: { col:14, row:9 },
          editAs: 'oneCell'
        });
      }

      // PENGATURAN CETAK A4
      ws.views = [{
        state: 'normal',
        showGridLines: false
      }];

      ws.pageSetup = {
        paperSize: 9,
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        printArea: 'A1:N38',
        margins: {
          left: 0.2,
          right: 0.2,
          top: 0.25,
          bottom: 0.25,
          header: 0.1,
          footer: 0.1
        }
      };

      return workbook;
    });
  }

  // ==========================================
  // 8. DOWNLOAD EXCEL
  // ==========================================

  async function downloadExcel() {
    if (!form.reportValidity()) return;

    if (typeof ExcelJS === 'undefined') {
      downloadStatus.textContent =
        'ExcelJS belum termuat. Periksa koneksi internet dan muat ulang halaman.';

      showToast('ExcelJS belum termuat.');

      return;
    }

    downloadButton.disabled = true;

    downloadButton.innerHTML =
      '<span>…</span> Menyusun dokumen';

    downloadStatus.textContent =
      'Memuat template Excel dan menyusun data...';

    try {
      const data = collectData();

      const response = await fetch('cv-template.xlsx', {
        cache: 'no-store'
      });

      if (!response.ok) {
        throw new Error(
          'Template Excel tidak ditemukan. Pastikan cv-template.xlsx satu folder dengan index.html.'
        );
      }

      data.templateBuffer = await response.arrayBuffer();

      const workbook = await makeExcelWorkbook(data);

      const buffer = await workbook.xlsx.writeBuffer();

      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      const url = URL.createObjectURL(blob);

      const link = document.createElement('a');

      link.href = url;

      link.download =
        `CV_Jepang_${cleanFilename(data.fullName)}.xlsx`;

      document.body.append(link);

      link.click();

      link.remove();

      window.setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 1500);

      downloadStatus.textContent =
        'CV Excel berhasil dibuat. Periksa kembali data, format, dan ejaan Katakana.';

      showToast('CV Jepang berhasil diunduh.');

    } catch (error) {
      console.error(error);

      downloadStatus.textContent =
        `Gagal membuat Excel: ${error.message || 'Kesalahan tidak diketahui.'}`;

      showToast('Gagal membuat dokumen Excel.');

    } finally {
      downloadButton.disabled = false;

      downloadButton.innerHTML =
        '<span>↓</span> Unduh CV Excel (.xlsx)';
    }
  }

  // ==========================================
  // 9. EVENT NAMA DAN FURIGANA
  // ==========================================

  $('#fullName').addEventListener('input', () => {
    if (!furiganaWasEdited) {
      $('#furigana').value = romajiToKatakana(value('fullName'));
    }

    renderPreview();
  });

  $('#furigana').addEventListener('input', () => {
    furiganaWasEdited = true;
    renderPreview();
  });

  $('#generateKana').addEventListener('click', () => {
    $('#furigana').value = romajiToKatakana(value('fullName'));

    furiganaWasEdited = false;

    renderPreview();
  });

  // ==========================================
  // 10. EVENT FOTO
  // ==========================================

  $('#photoInput').addEventListener('change', async event => {
    const file = event.target.files?.[0];

    if (!file) return;

    try {
      photoDataUrl = await cropPhoto(file);

      $('#photoPreview').innerHTML =
        `<img src="${photoDataUrl}" alt="Foto profil" />`;

      $('#removePhoto').hidden = false;

      renderPreview();

    } catch (error) {
      showToast(error.message || 'Foto tidak dapat diproses.');

      event.target.value = '';
    }
  });

  $('#removePhoto').addEventListener('click', () => {
    photoDataUrl = '';

    $('#photoInput').value = '';

    $('#photoPreview').innerHTML =
      '<span>写真</span><small>Foto 3 × 4</small>';

    $('#removePhoto').hidden = true;

    renderPreview();
  });

  // ==========================================
  // 11. PERNAH KE JEPANG
  // ==========================================

  $('#visitedJapan').addEventListener('change', () => {
    const enabled = value('visitedJapan') === 'yes';

    $('#travelPeriod').disabled = !enabled;
    $('#travelPurpose').disabled = !enabled;

    if (!enabled) {
      $('#travelPeriod').value = '';
      $('#travelPurpose').value = '';
    }

    renderPreview();
  });

  // ==========================================
  // 12. TAMBAH DAN HAPUS DATA
  // ==========================================

  document.addEventListener('click', event => {
    const add = event.target.closest('[data-add]');

    if (add) {
      addRow(add.dataset.add);
      return;
    }

    const remove = event.target.closest('.remove-row');

    if (remove) {
      const item = remove.closest('.repeat-item');

      if (!item) return;

      const type = item.dataset.rowType;

      item.remove();

      updateAddButton(type);

      renderPreview();
    }
  });

  // ==========================================
  // 13. SINKRONISASI FORM
  // ==========================================

  form.addEventListener('submit', event => {
    event.preventDefault();
  });

  form.addEventListener('input', renderPreview);

  form.addEventListener('change', renderPreview);

  downloadButton.addEventListener('click', downloadExcel);

  // ==========================================
  // 14. RESET FORM
  // ==========================================

  $('#resetForm').addEventListener('click', () => {
    if (!window.confirm('Hapus seluruh data yang sudah diisi?')) {
      return;
    }

    form.reset();

    Object.keys(containers).forEach(type => {
      containers[type].innerHTML = '';
      updateAddButton(type);
    });

    photoDataUrl = '';
    furiganaWasEdited = false;

    $('#photoInput').value = '';

    $('#photoPreview').innerHTML =
      '<span>写真</span><small>Foto 3 × 4</small>';

    $('#removePhoto').hidden = true;

    $('#ageValue').textContent = '—';

    $('#travelPeriod').disabled = true;
    $('#travelPurpose').disabled = true;

    downloadStatus.textContent = '';

    renderPreview();

    showToast('Formulir berhasil direset.');
  });

  // ==========================================
  // 15. INISIALISASI APLIKASI
  // ==========================================

  Object.keys(containers).forEach(updateAddButton);

  $('#travelPeriod').disabled = value('visitedJapan') !== 'yes';
  $('#travelPurpose').disabled = value('visitedJapan') !== 'yes';

  updateAgeDisplay();

  renderPreview();

})();
