// main.js - ETF 数据看板

document.addEventListener('DOMContentLoaded', () => {
    loadData();
});

function loadData() {
    // 尝试加载 Excel 数据（后续会根据实际数据格式调整）
    // fetch('assets/posts/data.json')
    //     .then(res => res.json())
    //     .then(data => renderData(data))
    //     .catch(err => {
    //         document.getElementById('data-container').innerHTML =
    //             '<p class="loading">暂无数据</p>';
    //     });

    // 更新最后更新时间
    const now = new Date();
    document.getElementById('last-update').textContent =
        now.toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });
}
