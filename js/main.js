/* ===== ETF 数据看板 · main.js ===== */

const POSTS = [];

const GALLERY = [];

const VIDEOS = [];

const MATERIALS = [
    {
        icon: "📊",
        title: "全市场数据分析表",
        desc: "深交所+上交所全部ETF数据，含份额、净值、市值、排行榜。",
        tags: ["ETF", "数据", "Excel", "全市场"],
        date: "2026-06-17",
        file: "assets/posts/全市场数据分析表02.xlsx",
        fileName: "全市场数据分析表02.xlsx"
    }
];

/* ===== 渲染 ===== */

function renderRecentPosts() {
    const container = document.getElementById("recent-posts");
    if (!container) return;

    if (MATERIALS.length === 0) {
        container.innerHTML = `<div class="empty-state">
            <div class="emoji">📝</div>
            <h3>还没有内容</h3>
            <p>数据采集完成后会自动更新。</p>
        </div>`;
        return;
    }

    container.innerHTML = MATERIALS.map(m => {
        return `<a class="post-card" href="materials.html">
            <div class="post-card-image" style="display:flex;align-items:center;justify-content:center;font-size:2.5rem;color:var(--text-secondary);background:var(--bg-alt)">${m.icon || '📄'}</div>
            <div class="post-card-body">
                <span class="post-card-tag">数据</span>
                <h3>${m.title}</h3>
                <p>${m.desc}</p>
                <div class="post-card-date">${m.date}</div>
            </div>
        </a>`;
    }).join("");
}

function renderMaterials() {
    const container = document.getElementById("materials-list");
    if (!container) return;

    if (MATERIALS.length === 0) {
        container.innerHTML = `<div class="empty-state">
            <div class="emoji">📚</div>
            <h3>还没有资料</h3>
            <p>数据采集完成后会自动更新。</p>
        </div>`;
        return;
    }

    container.innerHTML = MATERIALS.map(m => {
        const hasFile = m.file && m.fileName;
        const previewUrl = hasFile ? `preview.html?file=${encodeURIComponent(m.file)}&name=${encodeURIComponent(m.fileName)}` : '';
        const downloadUrl = hasFile ? m.file : '';

        let inner = `
            <div class="material-icon">${m.icon || '📄'}</div>
            <div class="material-content">
                <h3>${m.title}</h3>
                <p>${m.desc}</p>
                <div class="material-meta">
                    <span>${m.date}</span>
                    ${(m.tags || []).map(t => `<span class="material-tag">${t}</span>`).join('')}
                </div>
            </div>
        `;

        if (hasFile) {
            inner += `
                <div class="material-actions">
                    <a href="${previewUrl}" class="mat-action mat-preview">👁️ 预览</a>
                    <a href="${downloadUrl}" class="mat-action mat-dl" download onclick="event.stopPropagation()">⬇️</a>
                </div>
            `;
        }

        if (previewUrl) {
            return `<a class="material-item" href="${previewUrl}">${inner}</a>`;
        }
        return `<div class="material-item">${inner}</div>`;
    }).join("");
}

/* ===== 初始化 ===== */
document.addEventListener('DOMContentLoaded', () => {
    renderRecentPosts();
    renderMaterials();

    // 导航高亮
    const path = window.location.pathname;
    document.querySelectorAll('.nav-links a').forEach(a => {
        const href = a.getAttribute('href');
        if (href === 'index.html' && (path === '/' || path.endsWith('index.html'))) {
            a.classList.add('active');
        } else if (href !== 'index.html' && path.endsWith(href)) {
            a.classList.add('active');
        }
    });
});
