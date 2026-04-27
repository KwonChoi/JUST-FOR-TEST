// Main application logic
class TravelInfoApp {
    constructor() {
        this.countries = window.countriesData;
        this.currentFilter = 'all';
        this.searchTerm = '';
        this.init();
    }

    init() {
        this.renderCountries();
        this.bindEvents();
    }

    bindEvents() {
        // Search functionality
        const searchInput = document.getElementById('searchInput');
        const searchBtn = document.getElementById('searchBtn');

        searchBtn.addEventListener('click', () => this.handleSearch());
        searchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.handleSearch();
        });

        // Filter tabs
        document.querySelectorAll('.filter-tab').forEach(tab => {
            tab.addEventListener('click', (e) => this.handleFilter(e.target.dataset.filter));
        });

        // Modal close
        const modal = document.getElementById('countryModal');
        const closeBtn = document.querySelector('.close');
        closeBtn.addEventListener('click', () => this.closeModal());

        // Close modal when clicking outside
        modal.addEventListener('click', (e) => {
            if (e.target === modal) this.closeModal();
        });

        // Info tabs in modal
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('info-tab')) {
                this.handleInfoTabClick(e.target);
            }
        });
    }

    handleSearch() {
        const searchInput = document.getElementById('searchInput');
        this.searchTerm = searchInput.value.trim().toLowerCase();
        this.renderCountries();
    }

    handleFilter(filter) {
        this.currentFilter = filter;

        // Update active tab
        document.querySelectorAll('.filter-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.filter === filter);
        });

        this.renderCountries();
    }

    handleInfoTabClick(tab) {
        const section = tab.closest('.country-detail');
        const infoType = tab.dataset.info;

        // Update active tab
        section.querySelectorAll('.info-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        // Show corresponding content
        section.querySelectorAll('.info-section').forEach(section => section.classList.remove('active'));
        section.querySelector(`[data-section="${infoType}"]`).classList.add('active');
    }

    filterCountries() {
        return this.countries.filter(country => {
            const matchesFilter = this.currentFilter === 'all' || country.region === this.currentFilter;
            const matchesSearch = !this.searchTerm ||
                country.name.toLowerCase().includes(this.searchTerm) ||
                country.nameEn.toLowerCase().includes(this.searchTerm);
            return matchesFilter && matchesSearch;
        });
    }

    renderCountries() {
        const filteredCountries = this.filterCountries();
        const countryGrid = document.getElementById('countryGrid');

        if (filteredCountries.length === 0) {
            countryGrid.innerHTML = '<p style="text-align: center; grid-column: 1/-1; padding: 2rem;">没有找到符合条件的国家</p>';
            return;
        }

        countryGrid.innerHTML = filteredCountries.map(country => `
            <div class="country-card" onclick="app.showCountryDetail('${country.id}')">
                <img src="${country.flag}" alt="${country.name}国旗" class="country-flag">
                <div class="country-info">
                    <h3 class="country-name">${country.name}</h3>
                    <p class="country-region">${this.getRegionName(country.region)}</p>
                    <p class="country-highlight">${country.highlight}</p>
                </div>
            </div>
        `).join('');
    }

    getRegionName(region) {
        const regions = {
            'asia': '亚洲',
            'europe': '欧洲',
            'america': '美洲',
            'africa': '非洲',
            'oceania': '大洋洲'
        };
        return regions[region] || region;
    }

    showCountryDetail(countryId) {
        const country = this.countries.find(c => c.id === countryId);
        if (!country) return;

        const modal = document.getElementById('countryModal');
        const detailContainer = document.getElementById('countryDetail');

        detailContainer.innerHTML = `
            <div class="country-detail">
                <div class="country-detail-header">
                    <img src="${country.flag}" alt="${country.name}国旗" class="country-detail-flag">
                    <div class="country-detail-title">
                        <h2>${country.name}</h2>
                        <p>${country.nameEn} • ${country.capital} • ${country.population}人口</p>
                    </div>
                </div>

                <div class="info-tabs">
                    <button class="info-tab active" data-info="visa">签证信息</button>
                    <button class="info-tab" data-info="currency">货币汇率</button>
                    <button class="info-tab" data-info="language">语言交流</button>
                    <button class="info-tab" data-info="customs">习俗禁忌</button>
                    <button class="info-tab" data-info="routes">玩法路线</button>
                    <button class="info-tab" data-info="vocabulary">高频词汇</button>
                </div>

                <div class="info-content">
                    <!-- Visa Section -->
                    <div class="info-section active" data-section="visa">
                        <div class="info-item">
                            <h4>签证类型：${country.visa.type}</h4>
                            <p><strong>停留时间：</strong>${country.visa.duration}</p>
                            <p><strong>办理难度：</strong>${country.visa.difficulty}</p>
                            <p><strong>申请建议：</strong>${country.visa.tips}</p>
                        </div>
                        <div class="info-item">
                            <h4>申请材料：</h4>
                            <ul style="margin-left: 20px; line-height: 1.8;">
                                ${country.visa.requirements.map(req => `<li>${req}</li>`).join('')}
                            </ul>
                        </div>
                    </div>

                    <!-- Currency Section -->
                    <div class="info-section" data-section="currency">
                        <div class="info-item">
                            <h4>货币信息</h4>
                            <p><strong>货币名称：</strong>${country.currency.name} (${country.currency.code})</p>
                            <p><strong>货币符号：</strong>${country.currency.symbol}</p>
                            <p><strong>汇率参考：</strong>${country.currency.rate}</p>
                            <p><strong>支付建议：</strong>${country.currency.paymentTips}</p>
                        </div>
                    </div>

                    <!-- Language Section -->
                    <div class="info-section" data-section="language">
                        <div class="info-item">
                            <h4>官方语言：${country.language.official}</h4>
                        </div>
                        <div class="info-item">
                            <h4>常用问候语：</h4>
                            ${country.language.commonGreetings.map(greeting =>
                                `<p><strong>${greeting.chinese}：</strong>${greeting[Object.keys(greeting)[2]]}</p>`
                            ).join('')}
                        </div>
                        <div class="info-item">
                            <h4>实用短语：</h4>
                            ${country.language.usefulPhrases.map(phrase =>
                                `<p><strong>${phrase.chinese}：</strong>${phrase[Object.keys(phrase)[2]]}</p>`
                            ).join('')}
                        </div>
                    </div>

                    <!-- Customs Section -->
                    <div class="info-section" data-section="customs">
                        ${country.customs.map(custom => `
                            <div class="info-item">
                                <h4>${custom.title}</h4>
                                <p>${custom.content}</p>
                            </div>
                        `).join('')}
                    </div>

                    <!-- Routes Section -->
                    <div class="info-section" data-section="routes">
                        ${country.routes.map(route => `
                            <div class="route-item">
                                <span class="route-duration">${route.duration}</span>
                                <h4>${route.title}</h4>
                                <p>${route.description}</p>
                                <p><strong>主要景点：</strong>${route.highlights.join(' • ')}</p>
                            </div>
                        `).join('')}
                    </div>

                    <!-- Vocabulary Section -->
                    <div class="info-section" data-section="vocabulary">
                        <div class="vocabulary-grid">
                            ${country.vocabulary.map(vocab => `
                                <div class="vocab-item">
                                    <div class="vocab-chinese">${vocab.chinese}</div>
                                    <div class="vocab-english">${vocab.english}</div>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                </div>
            </div>
        `;

        modal.style.display = 'block';
        document.body.style.overflow = 'hidden';
    }

    closeModal() {
        const modal = document.getElementById('countryModal');
        modal.style.display = 'none';
        document.body.style.overflow = 'auto';
    }
}

// Initialize app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.app = new TravelInfoApp();
});

// Close modal with Escape key
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        const modal = document.getElementById('countryModal');
        if (modal.style.display === 'block') {
            window.app.closeModal();
        }
    }
});