/* ===== Mega Pro Suivi Tech V1 – App Logic ===== */

(function () {
    'use strict';

    // ─── Storage Keys ───
    const STORAGE_KEY = 'antigravity_clients';

    // ─── Utils ───
    const parseAmt = (val) => {
        if (typeof val === 'number') return val;
        if (typeof val === 'string') {
            return Number(val.replace(/[^0-9.-]/g, '')) || 0;
        }
        return 0;
    };
    const THEME_KEY = 'megapro_theme';

    // ─── Helpers ───
    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    function todayStr() {
        const d = new Date();
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    function normalizeDate(dStr) {
        if (!dStr) return '';
        // Format YYYY-MM-DD
        if (/^\d{4}-\d{2}-\d{2}$/.test(dStr)) return dStr;
        // Format DD/MM/YYYY ou D/M/YYYY
        if (dStr.includes('/')) {
            const parts = dStr.split('/');
            if (parts.length === 3) {
                // Ensure YYYY-MM-DD
                return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            }
        }
        return dStr;
    }

    function generateId() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
    }

    function formatMoney(n) {
        return Number(n || 0).toLocaleString('fr-FR');
    }

    // Version compatible jsPDF (espace normal au lieu de l'espace insécable)
    function formatMoneyPDF(n) {
        const num = Number(n || 0);
        return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    }

    /** Échappe les caractères spéciaux HTML pour éviter les injections XSS */
    function escapeHTML(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    /** Calcule le Total TTC : Facture + Frais Annexes + Déplacement */
    function calcTTC(c) {
        return (Number(c.montantBase) || 0) + (Number(c.fraisAnnexes) || 0) + (Number(c.deplacement) || 0);
    }

    function formatDateFR(iso) {
        const d = new Date(iso);
        return d.toLocaleDateString('fr-FR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    }

    function parseWhatsApp(text) {
        const lines = text.split('\n');
        const data = {
            codeClient: '',
            codeAgent: '',
            codeMarketeur: '',
            codeS: '',
            nomClient: '',
            telephone: '',
            intervention: '',
            montantBase: 0,
            commission: 0,
            fraisAnnexes: 0,
            deplacement: 0,
            commentaire: '',
            priorite: 'normal',
            statut: 'en_attente'
        };

        const cleanNum = (str) => {
            if (!str) return 0;
            return Number(str.replace(/[^0-9]/g, '')) || 0;
        };

        lines.forEach(line => {
            const l = line.toLowerCase();

            // Traitement spécial pour CODE S : on cherche prioritairement un pattern 000000-000000
            const codeSMatch = line.match(/\b\d{6}-\d{6}\b/);
            if (codeSMatch) {
                data.codeS = codeSMatch[0];
                return;
            } else if (l.includes('code s')) {
                const codeSParts = line.split(/[:]/);
                if (codeSParts.length >= 2) {
                    data.codeS = codeSParts.slice(1).join(':').trim();
                }
                return;
            }

            const parts = line.split(/[:=]/);
            if (parts.length < 2) return;
            const val = parts.slice(1).join(':').trim();
            if (!val) return;

            if (l.includes('code client') || l.includes('id client')) data.codeClient = val;
            else if (l.includes('code agent') || l.includes('id agent')) data.codeAgent = val;
            else if (l.includes('code mark') || l.includes('marketeur')) data.codeMarketeur = val;
            else if (l.includes('nom') || l.includes('prénom') || l.includes('prenom')) data.nomClient = val;
            else if (l.includes('tel') || l.includes('contact') || l.includes('tél')) data.telephone = val;
            else if (l.includes('inter') || l.includes('serv') || l.includes('prest')) data.intervention = val;
            else if ((l.includes('facture') && !l.includes('ttc')) || l.includes('montant base')) data.montantBase = cleanNum(val);
            else if (l.includes('commis')) data.commission = cleanNum(val);
            else if (l.includes('dépla') || l.includes('depla') || l.includes('trajet')) data.deplacement = cleanNum(val);
            else if (l.includes('frais') || l.includes('annexe')) data.fraisAnnexes = cleanNum(val);
            else if (l.includes('comment') || l.includes('note') || l.includes('obs')) data.commentaire = val;
            else if (l.includes('prior')) data.priorite = val.toLowerCase().includes('urg') ? 'urgent' : (val.toLowerCase().includes('moy') ? 'moyen' : 'normal');
        });

        return data;
    }

    // ─── Data Layer ───
    function getAllClients() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
        } catch {
            return [];
        }
    }

    function saveAllClients(clients) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(clients));
    }

    function getClientsForDate(dateStr) {
        const target = normalizeDate(dateStr);
        return getAllClients().filter(c => normalizeDate(c.date) === target);
    }

    function addClient(data) {
        const clients = getAllClients();
        const client = {
            id: generateId(),
            codeClient: data.codeClient || '',
            codeAgent: data.codeAgent || '',
            codeMarketeur: data.codeMarketeur || '',
            codeS: data.codeS || '',
            nomClient: data.nomClient || '',
            telephone: data.telephone || '',
            intervention: data.intervention || '',
            priorite: data.priorite || 'normal',
            statut: data.statut || 'en_attente',
            montantBase: Number(data.montantBase) || 0,
            commission: Number(data.commission) || 0,
            fraisAnnexes: Number(data.fraisAnnexes) || 0,
            deplacement: Number(data.deplacement) || 0,
            commentaire: data.commentaire || '',
            date: data.date || todayStr()
        };
        clients.push(client);
        saveAllClients(clients);
        return client;
    }


    // Fonction globale pour vérifier les doublons et ajouter
    function checkAndAddClient(data) {
        return new Promise((resolve) => {
            const clients = getAllClients();
            
            // Critères de doublon : même Code Client (si renseigné), ou même Trio (Téléphone + Intervention)
            const isDuplicate = clients.some(c => {
                const sameCode = data.codeClient && c.codeClient && data.codeClient.toLowerCase() === c.codeClient.toLowerCase();
                const sameContactService = data.telephone && c.telephone && 
                                          data.telephone.replace(/\s+/g, '') === c.telephone.replace(/\s+/g, '') &&
                                          data.intervention && c.intervention &&
                                          data.intervention.toLowerCase() === c.intervention.toLowerCase();
                return sameCode || sameContactService;
            });

            if (isDuplicate) {
                const modal = $('#global-duplicate-modal');
                const btnForce = $('#global-dup-force');
                const btnCancel = $('#global-dup-cancel');
                
                if (modal && btnForce && btnCancel) {
                    modal.classList.remove('hidden');
                    
                    const handleForce = () => {
                        cleanup();
                        resolve(addClient(data));
                    };
                    
                    const handleCancel = () => {
                        cleanup();
                        showToast('Ajout annulé (doublon détecté).', 'info');
                        resolve(null);
                    };
                    
                    const cleanup = () => {
                        modal.classList.add('hidden');
                        btnForce.removeEventListener('click', handleForce);
                        btnCancel.removeEventListener('click', handleCancel);
                    };
                    
                    btnForce.addEventListener('click', handleForce);
                    btnCancel.addEventListener('click', handleCancel);
                } else {
                    // Fallback si la modale n'est pas dispo
                    resolve(addClient(data));
                }
            } else {
                resolve(addClient(data));
            }
        });
    }

    function updateClient(id, updates) {
        const clients = getAllClients();
        const idx = clients.findIndex(c => c.id === id);
        if (idx === -1) return null;
        Object.assign(clients[idx], updates);
        saveAllClients(clients);
        return clients[idx];
    }

    function deleteClient(id) {
        const clients = getAllClients().filter(c => c.id !== id);
        saveAllClients(clients);
    }


    // ─── Theme Management ───
    function initTheme() {
        const savedTheme = localStorage.getItem(THEME_KEY) || 'dark';
        applyTheme(savedTheme);

        $$('.theme-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const theme = btn.getAttribute('data-theme');
                applyTheme(theme);
            });
        });
    }

    function applyTheme(theme) {
        document.body.classList.remove('theme-light', 'theme-cosmic');
        if (theme !== 'dark') {
            document.body.classList.add(`theme-${theme}`);
        }
        localStorage.setItem(THEME_KEY, theme);
        $$('.theme-btn').forEach(btn => {
            if (btn.getAttribute('data-theme') === theme) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });
    }

    // ─── Objective ───
    // Les objectifs sont calculés dynamiquement sur la base des fiches effectuées dans la période affichée.

    // ─── Priority & Status ───
    const PRIORITY_ORDER = { urgent: 0, moyen: 1, normal: 2 };
    const STATUS_LABELS = {
        en_attente: '⏳ En attente',
        en_cours: '🔧 En cours',
        effectue: '✅ Effectué',
        annule: '❌ Annulé'
    };

    function sortClients(clients) {
        return [...clients].sort((a, b) => (PRIORITY_ORDER[a.priorite] ?? 2) - (PRIORITY_ORDER[b.priorite] ?? 2));
    }

    function formatStatus(statut) {
        return STATUS_LABELS[statut] || statut;
    }

    // ─── UI State & Global Variables ───
    let currentScreen = 'dashboard';
    let viewDate = todayStr();
    let searchQuery = '';
    let filterStatus = 'all';
    let currentHistoryDate = null;

    // Chart instances
    let caChartInstance = null;
    let statusChartInstance = null;

    // ─── Render Logic ───
    function renderDashboard() {
        const start = $('#dash-start-date').value || todayStr();
        const end = $('#dash-end-date').value || todayStr();

        const clients = getAllClients().filter(c => c.date >= start && c.date <= end);
        const doneClients = clients.filter(c => c.statut === 'effectue');

        const totalExpected = clients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const totalRealised = doneClients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);

        const totalClients = clients.length;
        const doneCount = doneClients.length;

        const rate = totalClients > 0 ? Math.round((doneCount / totalClients) * 100) : 0;

        $('#kpi-total-clients').textContent = totalClients;
        $('#kpi-done').textContent = doneCount;
        $('#kpi-rate').textContent = rate + '%';
        if ($('#kpi-ca')) $('#kpi-ca').textContent = formatMoney(totalRealised);

        if ($('#progress-fill')) $('#progress-fill').style.width = Math.min(rate, 100) + '%';
        if ($('#progress-text')) $('#progress-text').textContent = `${doneCount} / ${totalClients} interventions réalisées`;
        if ($('#progress-pct')) $('#progress-pct').textContent = rate + '%';

        const counts = {
            en_attente: clients.filter(c => c.statut === 'en_attente').length,
            en_cours: clients.filter(c => c.statut === 'en_cours').length,
            effectue: doneClients.length,
            annule: clients.filter(c => c.statut === 'annule').length
        };
        const total = clients.length || 1;
        if ($('#stat-waiting')) $('#stat-waiting').style.width = (counts.en_attente / total * 100) + '%';
        if ($('#stat-progress')) $('#stat-progress').style.width = (counts.en_cours / total * 100) + '%';
        if ($('#stat-done')) $('#stat-done').style.width = (counts.effectue / total * 100) + '%';
        if ($('#stat-cancelled')) $('#stat-cancelled').style.width = (counts.annule / total * 100) + '%';

        if ($('#count-waiting')) $('#count-waiting').textContent = counts.en_attente;
        if ($('#count-progress')) $('#count-progress').textContent = counts.en_cours;
        if ($('#count-done')) $('#count-done').textContent = counts.effectue;
        if ($('#count-cancelled')) $('#count-cancelled').textContent = counts.annule;

        updateSidebarProgress();

        // ─── Financial Summary Cards (Dashboard) ─── combiné en une seule passe
        const totals = clients.reduce((acc, c) => {
            acc.factures += Number(c.montantBase) || 0;
            acc.annexes  += Number(c.fraisAnnexes) || 0;
            acc.depl     += Number(c.deplacement) || 0;
            acc.ttc      += calcTTC(c);
            return acc;
        }, { factures: 0, annexes: 0, depl: 0, ttc: 0 });

        if ($('#dash-total-factures'))    $('#dash-total-factures').textContent    = formatMoney(totals.factures);
        if ($('#dash-total-annexes'))     $('#dash-total-annexes').textContent     = formatMoney(totals.annexes);
        if ($('#dash-total-deplacement')) $('#dash-total-deplacement').textContent = formatMoney(totals.depl);
        if ($('#dash-total-ttc'))         $('#dash-total-ttc').textContent         = formatMoney(totals.ttc);
    }

    /** @deprecated Le widget sidebar a été supprimé. Cette fonction ne fait rien si les éléments sont absents. */
    function updateSidebarProgress(done, total) {
        const textEl = $('#sidebar-progress-text');
        const pctEl  = $('#sidebar-progress-pct');
        const barEl  = $('#sidebar-progress-bar');
        if (!textEl && !pctEl && !barEl) return; // Widget absent, sortie rapide

        if (done === undefined || total === undefined) {
            const start = ($('#dash-start-date') && $('#dash-start-date').value) || todayStr();
            const end   = ($('#dash-end-date')   && $('#dash-end-date').value)   || todayStr();
            const all   = getAllClients();
            const period = all.filter(c => c.date >= start && c.date <= end);
            done  = period.filter(c => c.statut === 'effectue').length;
            total = period.length;
        }
        const rate = total > 0 ? Math.round((done / total) * 100) : 0;
        if (textEl) textEl.textContent = `${done} / ${total}`;
        if (pctEl)  pctEl.textContent  = rate + '%';
        if (barEl)  barEl.style.width  = Math.min(rate, 100) + '%';
    }

    function renderClientList() {
        // Lire la plage de dates sélectionnée
        const start = ($('#clients-start-date') && $('#clients-start-date').value) || todayStr();
        const end = ($('#clients-end-date') && $('#clients-end-date').value) || todayStr();

        let clients = getAllClients().filter(c => c.date >= start && c.date <= end);

        if (filterStatus !== 'all') {
            clients = clients.filter(c => c.statut === filterStatus);
        }
        if (searchQuery) {
            const q = searchQuery.toLowerCase();
            clients = clients.filter(c =>
                (c.nomClient || '').toLowerCase().includes(q) ||
                (c.codeClient || '').toLowerCase().includes(q) ||
                (c.telephone || '').toLowerCase().includes(q) ||
                (c.codeMarketeur || '').toLowerCase().includes(q) ||
                (c.codeS || '').toLowerCase().includes(q)
            );
        }

        // ─── Financial Summary (Client List) ─── combiné en une seule passe
        const clTotals = clients.reduce((acc, c) => {
            acc.factures += Number(c.montantBase) || 0;
            acc.annexes  += Number(c.fraisAnnexes) || 0;
            acc.depl     += Number(c.deplacement) || 0;
            acc.ttc      += calcTTC(c);
            return acc;
        }, { factures: 0, annexes: 0, depl: 0, ttc: 0 });

        if ($('#cl-total-factures'))    $('#cl-total-factures').textContent    = formatMoney(clTotals.factures);
        if ($('#cl-total-annexes'))     $('#cl-total-annexes').textContent     = formatMoney(clTotals.annexes);
        if ($('#cl-total-deplacement')) $('#cl-total-deplacement').textContent = formatMoney(clTotals.depl);
        if ($('#cl-total-ttc'))         $('#cl-total-ttc').textContent         = formatMoney(clTotals.ttc);

        const doneInList = clients.filter(c => c.statut === 'effectue').length;
        const totalInList = clients.length;
        updateSidebarProgress(doneInList, totalInList);

        const tbody = $('#clients-tbody');
        tbody.innerHTML = '';

        if (clients.length === 0) {
            $('#empty-state').style.display = 'flex';
            $('#clients-table').style.display = 'none';
        } else {
            $('#empty-state').style.display = 'none';
            $('#clients-table').style.display = 'table';

            clients.forEach(c => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${escapeHTML(c.date || '-')}</td>
                    <td><span class="code-pill">${escapeHTML(c.codeClient || '-')}</span></td>
                    <td><span class="code-pill">${escapeHTML(c.codeAgent || '-')}</span></td>
                    <td>${escapeHTML(c.codeMarketeur || '-')}</td>
                    <td class="font-bold">${escapeHTML(c.nomClient || '-')}</td>
                    <td>${escapeHTML(c.telephone || '-')}</td>
                    <td>${escapeHTML(c.intervention || '-')}</td>
                    <td>${formatMoney(c.montantBase)}</td>
                    <td style="color: var(--warning);">${formatMoney(c.commission)}</td>
                    <td>${formatMoney(c.fraisAnnexes)}</td>
                    <td>${formatMoney(c.deplacement)}</td>
                    <td class="td-truncate" title="${escapeHTML(c.commentaire || '')}">${escapeHTML(c.commentaire || '-')}</td>
                    <td><span class="code-pill">${escapeHTML(c.codeS || '-')}</span></td>
                    <td><button class="status-badge status-${escapeHTML(c.statut)}" onclick="window.cycleStatus('${escapeHTML(c.id)}')">${formatStatus(c.statut)}</button></td>
                    <td>
                        <div class="action-btns">
                            <button class="action-btn" onclick="window.openEditModal('${escapeHTML(c.id)}')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button>
                            <button class="action-btn delete" onclick="window.askDelete('${escapeHTML(c.id)}')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg></button>
                        </div>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }
    }

    function renderHistory() {
        const start = $('#history-start-date').value;
        const end = $('#history-end-date').value;
        const clients = getAllClients();
        const datesMap = {};

        clients.forEach(c => {
            if ((!start || c.date >= start) && (!end || c.date <= end)) {
                if (!datesMap[c.date]) datesMap[c.date] = [];
                datesMap[c.date].push(c);
            }
        });

        const container = $('#history-container');
        if (!container) return;
        container.innerHTML = '';
        const sortedDates = Object.keys(datesMap).sort().reverse();

        if (sortedDates.length === 0) {
            const empty = $('#history-empty');
            if (empty) empty.style.display = 'flex';
        } else {
            const empty = $('#history-empty');
            if (empty) empty.style.display = 'none';
            sortedDates.forEach(date => {
                const dayClients = datesMap[date];
                const done = dayClients.filter(c => c.statut === 'effectue');
                const ca = dayClients.reduce((s, c) => s + (c.statut === 'effectue' ? (Number(c.montantBase) || 0) : 0), 0);

                const card = document.createElement('div');
                card.className = 'history-card';
                card.innerHTML = `
                    <div class="h-card-header">
                        <h3>${formatDateFR(date)}</h3>
                        <span class="h-card-ca">${formatMoney(ca)}</span>
                    </div>
                    <div class="h-card-stats">
                        <span>${dayClients.length} fiches</span>
                        <span>${done.length} effectuées</span>
                    </div>
                    <button class="btn btn-ghost" style="width:100%; margin-top:12px;" onclick="window.viewArchive('${date}')">Détails</button>
                `;
                container.appendChild(card);
            });
        }
    }

    function renderHistoryDetail(date) {
        if (!date) return;
        
        const clients = getAllClients().filter(c => c.date === date);
        const doneClients = clients.filter(c => c.statut === 'effectue');
        
        const totalExpected = clients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const totalRealised = doneClients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const rate = totalExpected > 0 ? Math.round((totalRealised / totalExpected) * 100) : 0;

        // Header Title
        if ($('#hd-date-title')) $('#hd-date-title').textContent = `Détails du ${formatDateFR(date)}`;

        // KPIs
        if ($('#hd-kpi-total-clients')) $('#hd-kpi-total-clients').textContent = clients.length;
        if ($('#hd-kpi-done')) $('#hd-kpi-done').textContent = doneClients.length;
        if ($('#hd-kpi-rate')) $('#hd-kpi-rate').textContent = rate + '%';
        if ($('#hd-kpi-ca')) $('#hd-kpi-ca').textContent = formatMoney(totalRealised);

        // Status Distribution
        const counts = {
            en_attente: clients.filter(c => c.statut === 'en_attente').length,
            en_cours: clients.filter(c => c.statut === 'en_cours').length,
            effectue: doneClients.length,
            annule: clients.filter(c => c.statut === 'annule').length
        };
        const total = clients.length || 1;
        if ($('#hd-stat-waiting')) $('#hd-stat-waiting').style.width = (counts.en_attente / total * 100) + '%';
        if ($('#hd-stat-progress')) $('#hd-stat-progress').style.width = (counts.en_cours / total * 100) + '%';
        if ($('#hd-stat-done')) $('#hd-stat-done').style.width = (counts.effectue / total * 100) + '%';
        if ($('#hd-stat-cancelled')) $('#hd-stat-cancelled').style.width = (counts.annule / total * 100) + '%';

        if ($('#hd-count-waiting')) $('#hd-count-waiting').textContent = counts.en_attente;
        if ($('#hd-count-progress')) $('#hd-count-progress').textContent = counts.en_cours;
        if ($('#hd-count-done')) $('#hd-count-done').textContent = counts.effectue;
        if ($('#hd-count-cancelled')) $('#hd-count-cancelled').textContent = counts.annule;

        // Financial Breakdown
        const totBase = clients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const totAnn = clients.reduce((s, r) => s + (Number(r.fraisAnnexes) || 0), 0);
        const totDepl = clients.reduce((s, r) => s + (Number(r.deplacement) || 0), 0);
        const totTTC = clients.reduce((s, r) => s + calcTTC(r), 0);

        if ($('#hd-total-base')) $('#hd-total-base').textContent = formatMoney(totBase);
        if ($('#hd-total-annexes')) $('#hd-total-annexes').textContent = formatMoney(totAnn);
        if ($('#hd-total-deplacement')) $('#hd-total-deplacement').textContent = formatMoney(totDepl);
        if ($('#hd-total-ttc')) $('#hd-total-ttc').textContent = formatMoney(totTTC);

        // Client Table
        const tbody = $('#hd-tbody');
        if (tbody) {
            tbody.innerHTML = clients.map(c => `
                <tr>
                    <td>${escapeHTML(c.date || '-')}</td>
                    <td><span class="code-pill">${escapeHTML(c.codeClient || '-')}</span></td>
                    <td><span class="code-pill">${escapeHTML(c.codeAgent || '-')}</span></td>
                    <td>${escapeHTML(c.codeMarketeur || '-')}</td>
                    <td class="font-bold">${escapeHTML(c.nomClient || '-')}</td>
                    <td>${escapeHTML(c.telephone || '-')}</td>
                    <td title="${escapeHTML(c.intervention || '')}">${c.intervention ? escapeHTML(c.intervention.length > 30 ? c.intervention.slice(0, 30) + '...' : c.intervention) : '-'}</td>
                    <td>${formatMoney(c.montantBase)}</td>
                    <td style="color: var(--warning);">${formatMoney(c.commission)}</td>
                    <td>${formatMoney(c.fraisAnnexes)}</td>
                    <td>${formatMoney(c.deplacement)}</td>
                    <td class="td-truncate" title="${escapeHTML(c.commentaire || '')}">${escapeHTML(c.commentaire || '-')}</td>
                    <td><span class="code-pill">${escapeHTML(c.codeS || '-')}</span></td>
                    <td><button class="status-badge status-${escapeHTML(c.statut)}" onclick="window.cycleStatus('${escapeHTML(c.id)}')">${formatStatus(c.statut)}</button></td>
                </tr>
            `).join('');
        }
    }

    function renderPerformance() {
        const start = $('#perf-start-date').value;
        const end = $('#perf-end-date').value;
        if (!start || !end) return;

        const all = getAllClients().filter(c => c.date >= start && c.date <= end);
        const done = all.filter(c => c.statut === 'effectue');
        const totalCA = all.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const realisedCA = done.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const rate = totalCA > 0 ? Math.round((realisedCA / totalCA) * 100) : 0;

        const lossAmount = all.reduce((s, c) => s + (c.statut !== 'effectue' && c.statut !== 'annule' ? (Number(c.montantBase) || 0) : 0), 0);
        const lossRate = totalCA > 0 ? Math.round((lossAmount / totalCA) * 100) : 0;

        // --- Basic Summary ---
        if ($('#perf-total-clients')) $('#perf-total-clients').textContent = all.length;
        if ($('#perf-done-clients')) $('#perf-done-clients').textContent = done.length;
        if ($('#perf-done-recap')) $('#perf-done-recap').textContent = done.length;
        if ($('#perf-expected-ca')) $('#perf-expected-ca').textContent = formatMoney(totalCA);
        if ($('#perf-realised-ca')) $('#perf-realised-ca').textContent = formatMoney(realisedCA);

        // --- Realization Progress ---
        if ($('#perf-rate-pct')) $('#perf-rate-pct').textContent = rate + '%';
        if ($('#perf-progress-fill')) $('#perf-progress-fill').style.width = rate + '%';

        // --- Loss Progress ---
        if ($('#perf-loss-rate')) $('#perf-loss-rate').textContent = lossRate + '%';
        if ($('#perf-loss-fill')) $('#perf-loss-fill').style.width = lossRate + '%';
        if ($('#perf-loss-amount')) $('#perf-loss-amount').textContent = formatMoney(lossAmount);
        if ($('#perf-loss-count')) $('#perf-loss-count').textContent = all.filter(c => c.statut !== 'effectue' && c.statut !== 'annule').length;

        // --- Performance Note /20 ---
        const score = Math.round((rate / 100) * 20);
        if ($('#perf-score')) $('#perf-score').textContent = score;

        // Radial Gauge Animation
        const circle = $('#perf-score-circle');
        if (circle) {
            const pct = (score / 20) * 100;
            circle.setAttribute('stroke-dasharray', `${pct} 100`);
        }

        // Performance Label
        const label = $('#perf-score-label');
        if (label) {
            if (score >= 18) { label.textContent = 'Excellent'; label.style.color = '#22c55e'; }
            else if (score >= 14) { label.textContent = 'Très Bien'; label.style.color = '#10b981'; }
            else if (score >= 10) { label.textContent = 'Bien'; label.style.color = '#3b82f6'; }
            else if (score >= 6) { label.textContent = 'Passable'; label.style.color = '#eab308'; }
            else { label.textContent = 'Insuffisant'; label.style.color = '#ef4444'; }
        }

        // Score Breakdown
        if ($('#perf-score-fiches')) $('#perf-score-fiches').textContent = `${done.length} / ${all.length}`;
        if ($('#perf-score-ca')) $('#perf-score-ca').textContent = `${formatMoney(realisedCA)} FG`;
        if ($('#perf-score-loss')) {
            const lossFiches = all.filter(c => c.statut !== 'effectue' && c.statut !== 'annule').length;
            $('#perf-score-loss').textContent = `${lossFiches} fiches`;
        }

        // --- Charts ---
        renderCharts(all, start, end);
    }

    function renderCharts(clients, start, end) {
        // Destroy existing charts if any
        if (caChartInstance) caChartInstance.destroy();
        if (statusChartInstance) statusChartInstance.destroy();

        // --- 1. Line Chart (Évolution CA) ---
        // Group by date
        const caByDate = {};
        let currDate = new Date(start);
        const endDate = new Date(end);

        // Initialize all dates in range with 0 to ensure continuity
        while (currDate <= endDate) {
            const dateStr = currDate.toISOString().slice(0, 10);
            caByDate[dateStr] = 0;
            currDate.setDate(currDate.getDate() + 1);
        }

        // Fill actual data (only for 'effectue')
        clients.forEach(c => {
            if (c.statut === 'effectue' && caByDate[c.date] !== undefined) {
                caByDate[c.date] += (Number(c.montantBase) || 0);
            }
        });

        const labels = Object.keys(caByDate).sort();
        const data = labels.map(l => caByDate[l]);
        // Format labels for display (e.g. "05/10")
        const displayLabels = labels.map(l => {
            const d = new Date(l);
            return `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}`;
        });

        const ctxCa = $('#ca-chart');
        if (ctxCa) {
            Chart.defaults.color = '#94a3b8'; // text-muted
            Chart.defaults.font.family = "'Inter', sans-serif";

            caChartInstance = new Chart(ctxCa, {
                type: 'line',
                data: {
                    labels: displayLabels,
                    datasets: [{
                        label: 'CA Réalisé',
                        data: data,
                        borderColor: '#10b981', // success
                        backgroundColor: 'rgba(16, 185, 129, 0.1)',
                        borderWidth: 2,
                        pointBackgroundColor: '#10b981',
                        pointBorderColor: '#fff',
                        pointRadius: 4,
                        fill: true,
                        tension: 0.3
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    return formatMoney(context.raw);
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255, 255, 255, 0.05)' },
                            ticks: {
                                callback: function (value) {
                                    if (value === 0) return 0;
                                    return (value / 1000) + 'k';
                                }
                            }
                        },
                        x: {
                            grid: { display: false }
                        }
                    }
                }
            });
        }

        // --- 2. Doughnut Chart (Répartition Statuts) ---
        const counts = {
            en_attente: clients.filter(c => c.statut === 'en_attente').length,
            en_cours: clients.filter(c => c.statut === 'en_cours').length,
            effectue: clients.filter(c => c.statut === 'effectue').length,
            annule: clients.filter(c => c.statut === 'annule').length
        };

        const ctxStatus = $('#status-chart');
        if (ctxStatus) {
            statusChartInstance = new Chart(ctxStatus, {
                type: 'doughnut',
                data: {
                    labels: ['En attente', 'En cours', 'Effectué', 'Annulé'],
                    datasets: [{
                        data: [counts.en_attente, counts.en_cours, counts.effectue, counts.annule],
                        backgroundColor: [
                            '#f59e0b', // waiting
                            '#3b82f6', // progress
                            '#10b981', // done
                            '#ef4444'  // cancelled
                        ],
                        borderWidth: 0,
                        hoverOffset: 4
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '70%',
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                usePointStyle: true,
                                padding: 15,
                                font: { size: 11 }
                            }
                        }
                    }
                }
            });
        }
    }

    // ─── Control Logic ───
    function switchScreen(id) {
        // --- Security Check for Extraction ---
        if (id === 'extraction') {
            openExtractionSecurityModal();
            return; // Don't switch screen yet
        }

        currentScreen = id;
        $$('.screen').forEach(s => s.classList.remove('active'));
        $(`#screen-${id}`).classList.add('active');
        $$('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.screen === id));
        refresh();
    }

    function openExtractionSecurityModal() {
        const modal = $('#xt-security-modal');
        const keyInput = $('#xt-security-key');
        const errorMsg = $('#xt-security-error');
        
        if (!modal || !keyInput) return;
        
        // Reset modal state
        keyInput.value = '';
        if (errorMsg) errorMsg.style.display = 'none';
        
        modal.classList.remove('hidden');
        keyInput.focus();
    }

    function validateExtractionKey() {
        const keyInput = $('#xt-security-key');
        const errorMsg = $('#xt-security-error');
        const modal = $('#xt-security-modal');
        
        if (!keyInput) return;
        
        if (keyInput.value === 'Megapro') {
            modal.classList.add('hidden');
            // Now actually switch to extraction screen
            currentScreen = 'extraction';
            $$('.screen').forEach(s => s.classList.remove('active'));
            $('#screen-extraction').classList.add('active');
            $$('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.screen === 'extraction'));
            refresh();
            showToast('Accès autorisé ✅', 'success');
        } else {
            if (errorMsg) errorMsg.style.display = 'block';
            keyInput.style.borderColor = 'var(--danger)';
            setTimeout(() => keyInput.style.borderColor = 'var(--glass-border)', 1000);
        }
    }

    function refresh() {
        if (currentScreen === 'dashboard') renderDashboard();
        if (currentScreen === 'clients') renderClientList();
        if (currentScreen === 'history') renderHistory();
        if (currentScreen === 'history-detail') renderHistoryDetail(currentHistoryDate);
        if (currentScreen === 'performance') renderPerformance();
        // Extraction screen doesn't need a dynamic render function on load
    }

    // ─── Event Handlers ───
    function init() {
        initTheme();

        // Synchronisation du selecteur de date de l'écran d'ajout avec le champ du formulaire
        const addScreenDate = $('#add-screen-date');
        if (addScreenDate) {
            addScreenDate.addEventListener('change', () => {
                const newDate = addScreenDate.value;
                if (newDate) {
                    viewDate = newDate;
                    if ($('#input-date')) $('#input-date').value = newDate;
                    // On met aussi à jour le sélecteur de la vue clients au cas où
                    if ($('#clients-view-date')) $('#clients-view-date').value = newDate;
                }
            });
        }

        if ($('#dash-start-date')) $('#dash-start-date').value = todayStr();
        if ($('#dash-end-date')) $('#dash-end-date').value = todayStr();

        // Setup default dates for performance (Month start to Today)
        if ($('#perf-start-date')) {
            const d = new Date();
            const startMonth = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
            $('#perf-start-date').value = startMonth;
        }
        if ($('#perf-end-date')) $('#perf-end-date').value = todayStr();

        // Setup default dates for extraction
        if ($('#extract-start-date')) {
            const d = new Date();
            const startMonth = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
            $('#extract-start-date').value = startMonth;
        }
        if ($('#extract-end-date')) $('#extract-end-date').value = todayStr();

        // ─── Clients Date Filter ───
        const clientsViewDate = $('#clients-view-date');
        if (clientsViewDate) {
            clientsViewDate.addEventListener('change', (e) => {
                const newDate = e.target.value;
                if (newDate) {
                    viewDate = newDate;
                    renderClientList();
                }
            });
        }

        // ─── Dashboard Filter ───
        const btnDashFilter = $('#btn-dash-filter');
        if (btnDashFilter) {
            btnDashFilter.addEventListener('click', () => {
                renderDashboard();
                showToast('Filtres du tableau de bord appliqués', 'success');
            });
        }

        // Objectif manuel retiré : La logique est maintenant basée sur le comptage des fiches
        const objectiveInput = $('#objective-input');
        if (objectiveInput) {
            objectiveInput.parentElement.style.display = 'none';
        }
        const btnDashToday = $('#btn-dash-today');
        if (btnDashToday) {
            btnDashToday.addEventListener('click', () => {
                $('#dash-start-date').value = todayStr();
                $('#dash-end-date').value = todayStr();
                renderDashboard();
                showToast("Retour à la date d'aujourd'hui", 'info');
            });
        }

        // ─── Client List Date Range Filter ───
        if ($('#clients-start-date')) $('#clients-start-date').value = todayStr();
        if ($('#clients-end-date')) $('#clients-end-date').value = todayStr();

        const btnClientsApply = $('#btn-clients-filter-apply');
        if (btnClientsApply) {
            btnClientsApply.addEventListener('click', () => {
                renderClientList();
                showToast('Filtres appliqués', 'success');
            });
        }

        const btnResetPeriod = $('#btn-reset-period');
        if (btnResetPeriod) {
            btnResetPeriod.addEventListener('click', () => {
                if ($('#clients-start-date')) $('#clients-start-date').value = todayStr();
                if ($('#clients-end-date')) $('#clients-end-date').value = todayStr();
                renderClientList();
                showToast("Période réinitialisée à aujourd'hui", 'info');
            });
        }

        $('#mobile-menu-btn').addEventListener('click', () => {
            $('#sidebar').classList.toggle('open');
            $('#sidebar-overlay').classList.toggle('show');
        });

        $('#sidebar-overlay').addEventListener('click', () => {
            $('#sidebar').classList.remove('open');
            $('#sidebar-overlay').classList.remove('show');
        });

        // ─── Modal À propos ───
        const aboutModal = $('#about-modal');
        const btnAbout = $('#btn-about');
        const btnAboutClose = $('#btn-about-close');

        if (btnAbout && aboutModal) {
            btnAbout.addEventListener('click', () => {
                aboutModal.classList.remove('hidden');
                // Fermer la sidebar mobile si elle est ouverte
                $('#sidebar').classList.remove('open');
                $('#sidebar-overlay').classList.remove('show');
            });
        }

        if (btnAboutClose && aboutModal) {
            btnAboutClose.addEventListener('click', () => {
                aboutModal.classList.add('hidden');
            });
        }

        if (aboutModal) {
            // Fermer au clic sur l'overlay
            aboutModal.addEventListener('click', (e) => {
                if (e.target === aboutModal) {
                    aboutModal.classList.add('hidden');
                }
            });
            // Fermer avec Echap
            document.addEventListener('keydown', (e) => {
                if (e.key === 'Escape' && !aboutModal.classList.contains('hidden')) {
                    aboutModal.classList.add('hidden');
                }
            });
        }

        // ─── Toggle Add Client Form ───
        const btnToggleForm = $('#btn-toggle-add-form');
        const btnEmptyAdd   = $('#btn-empty-add');
        const addContainer = $('#add-client-container');

        const openAddContainer = () => {
            addContainer.classList.remove('hidden');
            if (btnToggleForm) btnToggleForm.classList.add('btn-ghost');
            if (!addContainer.classList.contains('hidden')) {
                $('#add-screen-date').value = viewDate;
                $('#input-date').value = viewDate;
                addContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        };

        if (btnToggleForm && addContainer) {
            btnToggleForm.addEventListener('click', () => {
                const isHidden = addContainer.classList.contains('hidden');
                if (isHidden) {
                    openAddContainer();
                } else {
                    addContainer.classList.add('hidden');
                    btnToggleForm.classList.remove('btn-ghost');
                }
            });
        }

        if (btnEmptyAdd && addContainer) {
            btnEmptyAdd.addEventListener('click', openAddContainer);
        }

        // ─── Export Dropdown ───
        $('#btn-export-toggle').addEventListener('click', (e) => {
            e.stopPropagation();
            $('#export-menu').classList.toggle('hidden');
        });
        document.addEventListener('click', () => {
            $('#export-menu').classList.add('hidden');
        });
        $('#btn-export-csv').addEventListener('click', () => { $('#export-menu').classList.add('hidden'); exportToCSV(); });
        $('#btn-export-excel').addEventListener('click', () => { $('#export-menu').classList.add('hidden'); exportToExcel(); });
        $('#btn-export-pdf').addEventListener('click', () => { $('#export-menu').classList.add('hidden'); exportToPDF(); });
        $('#btn-export-json').addEventListener('click', () => { $('#export-menu').classList.add('hidden'); exportToJSON(); });

        // ─── Export Performance PDF ───
        const btnExportPerfPdf = $('#btn-export-perf-pdf');
        if (btnExportPerfPdf) {
            btnExportPerfPdf.addEventListener('click', () => {
                exportPerformancePDF();
            });
        }

        updateSidebarProgress();

        // ─── Import Data (Multi-format) ───

        // Toggle import dropdown
        const btnImportToggle = $('#btn-import-toggle');
        if (btnImportToggle) {
            btnImportToggle.addEventListener('click', (e) => {
                e.stopPropagation();
                $('#import-menu').classList.toggle('hidden');
                // Close export menu if open
                $('#export-menu').classList.add('hidden');
            });
        }
        // Close import menu on document click
        document.addEventListener('click', () => {
            const im = $('#import-menu');
            if (im) im.classList.add('hidden');
        });

        // Trigger file inputs from dropdown buttons
        const btnImportExcel = $('#btn-import-excel');
        if (btnImportExcel) btnImportExcel.addEventListener('click', () => { $('#import-menu').classList.add('hidden'); $('#input-import-excel').click(); });

        const btnImportCsv = $('#btn-import-csv');
        if (btnImportCsv) btnImportCsv.addEventListener('click', () => { $('#import-menu').classList.add('hidden'); $('#input-import-csv').click(); });

        const btnImportPdf = $('#btn-import-pdf');
        if (btnImportPdf) btnImportPdf.addEventListener('click', () => { $('#import-menu').classList.add('hidden'); $('#input-import-pdf').click(); });

        // Restore from JSON (inside import dropdown)
        const btnJsonRestore = $('#btn-import-json-restore');
        if (btnJsonRestore) btnJsonRestore.addEventListener('click', () => { $('#import-menu').classList.add('hidden'); $('#input-import-json').click(); });


        // ── Excel import handler ──
        $('#input-import-excel').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            importFromSpreadsheet(file, 'Excel');
            e.target.value = '';
        });

        // ── CSV import handler ──
        $('#input-import-csv').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            importFromSpreadsheet(file, 'CSV');
            e.target.value = '';
        });

        // ── PDF import handler ──
        $('#input-import-pdf').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            importFromPDF(file);
            e.target.value = '';
        });

        // ── JSON backup import handler ──
        $('#input-import-json').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function (event) {
                try {
                    const importedData = JSON.parse(event.target.result);
                    if (!Array.isArray(importedData)) {
                        throw new Error('Format invalide : le fichier doit contenir un tableau de clients.');
                    }

                    const isValid = importedData.every(c => c.id && c.date !== undefined);
                    if (!isValid) {
                        throw new Error('Les données du fichier ne correspondent pas au format attendu.');
                    }

                    saveAllClients(importedData);
                    refresh();
                    showToast('Restauration réussie ! ✅');
                } catch (err) {
                    console.error(err);
                    showToast('Erreur lors de la lecture du fichier de sauvegarde.', 'error');
                }
            };
            reader.readAsText(file);
            e.target.value = '';
        });


        // ─── Reset Data ───
        $('#btn-reset-data').addEventListener('click', () => {
            $('#reset-confirm-input').value = '';
            $('#reset-error').style.display = 'none';
            $('#reset-modal').classList.remove('hidden');
        });
        $('#reset-cancel').addEventListener('click', () => {
            $('#reset-modal').classList.add('hidden');
        });
        $('#reset-execute').addEventListener('click', () => {
            const val = $('#reset-confirm-input').value.trim();
            if (val === 'YRA') {
                localStorage.removeItem(STORAGE_KEY);
                $('#reset-modal').classList.add('hidden');
                refresh();
                updateSidebarProgress();
                showToast('Toutes les données ont été supprimées.', 'info');
            } else {
                $('#reset-error').style.display = 'block';
            }
        });

        $('#search-input').addEventListener('input', (e) => {
            searchQuery = e.target.value;
            renderClientList();
        });

        // ─── Performance Filter ───
        const btnPerfFilter = $('#btn-perf-filter');
        if (btnPerfFilter) {
            btnPerfFilter.addEventListener('click', () => {
                renderPerformance();
                showToast('Filtres appliqués', 'success');
            });
        }

        // Navigation (Sidebar)
        $$('.nav-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const target = e.currentTarget.dataset.screen;
                // Pré-remplir la date si on va vers le formulaire d'ajout
                // Pré-remplir la date si on va vers le module de gestion des clients
                if (target === 'clients') {
                    if ($('#input-date')) $('#input-date').value = viewDate;
                    if ($('#add-screen-date')) $('#add-screen-date').value = viewDate;
                }
                switchScreen(target);
                if (window.innerWidth <= 768) {
                    $('#sidebar').classList.remove('open');
                    $('#sidebar-overlay').classList.remove('show'); // Changed from 'open' to 'show' to match existing overlay behavior
                }
            });
        });

        const btnHDBack = $('#btn-hd-back');
        if (btnHDBack) {
            btnHDBack.addEventListener('click', () => {
                switchScreen('history');
            });
        }

        // ================= Outil d'Extraction Logic =================
        const btnExtractExcel = $('#btn-extract-excel');
        if (btnExtractExcel) {
            btnExtractExcel.addEventListener('click', () => {
                const sDate = $('#extract-start-date').value;
                const eDate = $('#extract-end-date').value;
                
                if (!sDate || !eDate) {
                    showToast('Veuillez sélectionner une plage de dates valide.', 'error');
                    return;
                }
                
                const clientsToExport = getAllClients().filter(c => c.date >= sDate && c.date <= eDate);
                
                if (clientsToExport.length === 0) {
                    showToast('Aucune donnée à exporter pour cette période.', 'error');
                    return;
                }
                
                exportToCSV(clientsToExport, `Export_Clients_${sDate}_au_${eDate}.csv`);
                showToast('Export Excel/CSV généré avec succès !', 'success');
            });
        }

        const btnExtractPdf = $('#btn-extract-pdf');
        if (btnExtractPdf) {
            btnExtractPdf.addEventListener('click', () => {
                const sDate = $('#extract-start-date').value;
                const eDate = $('#extract-end-date').value;
                
                if (!sDate || !eDate) {
                    showToast('Veuillez sélectionner une plage de dates valide.', 'error');
                    return;
                }
                
                const clientsToExport = getAllClients().filter(c => c.date >= sDate && c.date <= eDate);
                
                if (clientsToExport.length === 0) {
                    showToast('Aucune donnée à exporter pour cette période.', 'error');
                    return;
                }
                
                // Assuming simple raw print-to-pdf using a temporary hidden table
                const printWindow = window.open('', '', 'width=800,height=600');
                if (!printWindow) {
                    showToast('Veuillez autoriser les pop-ups pour générer le PDF.', 'error');
                    return;
                }
                
                let html = `<html><head><title>Bilan des interventions du ${sDate} au ${eDate}</title>
                    <style>
                        body { font-family: sans-serif; padding: 20px; }
                        h1 { color: #1e40af; }
                        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                        th, td { border: 1px solid #cbd5e1; padding: 10px; text-align: left; }
                        th { background: #f1f5f9; }
                    </style>
                </head><body>
                    <h1>Bilan des interventions - Mega Pro</h1>
                    <h3>Période : ${sDate} au ${eDate}</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>Date</th>
                                <th>Code Client</th>
                                <th>Service</th>
                                <th>Montant TTC</th>
                                <th>Statut</th>
                            </tr>
                        </thead>
                        <tbody>`;
                
                clientsToExport.forEach(c => {
                    html += `<tr>
                        <td>${c.date}</td>
                        <td>${c.codeClient || ''}</td>
                        <td>${c.intervention || ''}</td>
                        <td>${formatMoney(c.montantBase || 0)}</td>
                        <td>${c.statut || ''}</td>
                    </tr>`;
                });
                
                html += `</tbody></table></body></html>`;
                
                printWindow.document.write(html);
                printWindow.document.close();
                printWindow.onload = function() {
                    printWindow.print();
                    showToast('Génération du bilan PDF...', 'success');
                };
            });
        }

        const btnClearAll = $('#btn-clear-all');
        if (btnClearAll) {
            btnClearAll.addEventListener('click', () => {
                openClearAllModal();
            });
        }

        function openClearAllModal() {
            // Re-using the prompt-modal structure for the clear all confirmation
            const modal = document.createElement('div');
            modal.className = 'modal-overlay';
            modal.innerHTML = `
                <div class="modal" style="border-top: 4px solid #ef4444; max-width: 450px;">
                    <div class="modal-header">
                        <h2 style="color: #ef4444; display: flex; align-items: center; gap: 8px;">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width: 24px; height: 24px;">
                                <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                                <line x1="12" y1="9" x2="12" y2="13"/>
                                <line x1="12" y1="17" x2="12.01" y2="17"/>
                            </svg>
                            Action Irréversible
                        </h2>
                    </div>
                    <div style="padding: 0 24px 24px;">
                        <p style="margin-bottom: 20px; color: var(--text-primary);">Vous êtes sur le point de supprimer l'intégralité des fiches clients. Cette action ne peut pas être annulée.</p>
                        <p style="margin-bottom: 12px; font-weight: 500;">Pour confirmer, tapez <strong>YRA</strong> ci-dessous :</p>
                        <input type="text" id="clear-all-confirm-input" placeholder="YRA" autocomplete="off" style="width: 100%; padding: 12px; border: 1px solid var(--glass-border); background: var(--bg); color: var(--text-primary); border-radius: 8px; font-size: 1.1rem; text-align: center; font-weight: bold;">
                    </div>
                    <div class="modal-actions" style="border-top: 1px solid var(--glass-border); display: flex; justify-content: flex-end; gap: 12px; padding: 16px 24px; background: rgba(0,0,0,0.1);">
                        <button class="btn btn-secondary btn-cancel-clear" style="margin: 0;">Fermer</button>
                        <button class="btn btn-primary btn-confirm-clear" style="margin: 0; background: #ef4444; border-color: #ef4444;" disabled>Supprimer la base</button>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);

            const btnCancel = modal.querySelector('.btn-cancel-clear');
            const btnConfirm = modal.querySelector('.btn-confirm-clear');
            const input = modal.querySelector('#clear-all-confirm-input');

            // Handle input change to enable the delete button
            input.addEventListener('input', (e) => {
                if (e.target.value.trim() === 'YRA') {
                    btnConfirm.disabled = false;
                    btnConfirm.style.opacity = '1';
                } else {
                    btnConfirm.disabled = true;
                    btnConfirm.style.opacity = '0.5';
                }
            });

            btnCancel.addEventListener('click', () => modal.remove());
            
            btnConfirm.addEventListener('click', () => {
                if (input.value.trim() === 'YRA') {
                    localStorage.removeItem(STORAGE_KEY);
                    modal.remove();
                    showToast('La base de données a été purgée avec succès.', 'success');
                    refresh();
                }
            });
        }

        // ─── Status Filters in Client List ───
        $$('.filter-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                $$('.filter-btn').forEach(b => b.classList.remove('active'));
                const clickedBtn = e.currentTarget;
                clickedBtn.classList.add('active');
                filterStatus = clickedBtn.dataset.filter;
                renderClientList();
            });
        });

        // ─── Clear Current List ───
        const btnClearList = $('#btn-clear-current-list');
        if (btnClearList) {
            btnClearList.addEventListener('click', () => {
                const start = ($('#clients-start-date') && $('#clients-start-date').value) || todayStr();
                const end = ($('#clients-end-date') && $('#clients-end-date').value) || todayStr();

                let currentFiltered = getAllClients().filter(c => c.date >= start && c.date <= end);
                if (filterStatus !== 'all') {
                    currentFiltered = currentFiltered.filter(c => c.statut === filterStatus);
                }
                if (searchQuery) {
                    const q = searchQuery.toLowerCase();
                    currentFiltered = currentFiltered.filter(c =>
                        (c.nomClient || '').toLowerCase().includes(q) ||
                        (c.codeClient || '').toLowerCase().includes(q) ||
                        (c.telephone || '').toLowerCase().includes(q) ||
                        (c.codeMarketeur || '').toLowerCase().includes(q)
                    );
                }

                if (currentFiltered.length === 0) return;

                const idsToRemove = currentFiltered.map(c => c.id);
                let allClients = getAllClients().filter(c => !idsToRemove.includes(c.id));
                localStorage.setItem(STORAGE_KEY, JSON.stringify(allClients));
                
                showToast(idsToRemove.length + ' fiche(s) supprimée(s)', 'info');
                renderClientList();
                renderDashboard();
                renderHistory();
                renderPerformance();
            });
        }

        $('#client-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const data = {
                codeClient: $('#input-code-client').value.trim(),
                codeAgent: $('#input-code-agent').value.trim(),
                codeMarketeur: $('#input-code-marketeur').value.trim(),
                nomClient: $('#input-nom-client').value.trim(),
                telephone: $('#input-telephone').value.trim(),
                intervention: $('#input-intervention').value.trim(),
                priorite: $('#input-priorite') ? $('#input-priorite').value : 'normal',
                statut: $('#input-statut') ? $('#input-statut').value : 'en_attente',
                montantBase: $('#input-montant-base').value,
                commission: $('#input-commission').value,
                fraisAnnexes: $('#input-montant-annexes').value,
                deplacement: $('#input-montant-deplacement').value,
                commentaire: $('#input-commentaire').value.trim(),
                codeS: $('#input-code-s') ? $('#input-code-s').value.trim() : '',
                date: $('#input-date') ? $('#input-date').value : $('#add-screen-date').value || todayStr()
            };

            if (!data.telephone || !data.intervention) {
                showToast('Veuillez remplir les champs obligatoires', 'error');
                return;
            }

            const added = await checkAndAddClient(data);
            if (added) {
                showToast('Nouveau client enregistré ! ✅');
                e.target.reset();

                // Synchroniser les filtres de date sur la date du client ajouté
                viewDate = data.date;
                if ($('#clients-start-date')) $('#clients-start-date').value = data.date;
                if ($('#clients-end-date')) $('#clients-end-date').value = data.date;
                
                // Réinitialiser les filtres de statut et de recherche pour garantir l'affichage
                filterStatus = 'all';
                searchQuery = '';
                if ($('#search-input')) $('#search-input').value = '';
                $$('.filter-btn').forEach(b => b.classList.toggle('active', b.dataset.filter === 'all'));
                
                switchScreen('clients');
            }
        });

        $('#edit-form').addEventListener('submit', (e) => {
            e.preventDefault();
            const id = $('#edit-id').value;
            const clients = getAllClients();
            const idx = clients.findIndex(c => c.id === id);
            if (idx > -1) {
                // Merge old data with new submitted data
                const oldClient = clients[idx];
                const updatedClient = {
                    ...oldClient,
                    codeClient: $('#edit-code-client').value.trim(),
                    codeAgent: $('#edit-code-agent').value.trim(),
                    codeMarketeur: $('#edit-code-marketeur').value.trim(),
                    nomClient: $('#edit-nom-client').value.trim(),
                    telephone: $('#edit-telephone').value.trim(),
                    intervention: $('#edit-intervention').value.trim(),
                    priorite: $('#edit-priorite') ? $('#edit-priorite').value : oldClient.priorite,
                    statut: $('#edit-statut') ? $('#edit-statut').value : oldClient.statut,
                    montantBase: Number($('#edit-montant-base').value) || 0,
                    commission: Number($('#edit-commission').value) || 0,
                    fraisAnnexes: Number($('#edit-montant-annexes').value) || 0,
                    deplacement: Number($('#edit-montant-deplacement').value) || 0,
                    commentaire: $('#edit-commentaire').value.trim(),
                    codeS: $('#edit-code-s') ? $('#edit-code-s').value.trim() : ''
                };

                clients[idx] = updatedClient;
                saveAllClients(clients);
                showToast('Modifications enregistrées ✅');
                window.closeEditModal();
                refresh();
            }
        });

        $('#btn-cancel-edit').addEventListener('click', window.closeEditModal);
        $('#modal-close-edit').addEventListener('click', window.closeEditModal);

        $('#confirm-cancel').addEventListener('click', () => {
            $('#confirm-modal').classList.add('hidden');
            window.deleteTargetId = null;
        });

        $('#confirm-delete').addEventListener('click', () => {
            if (window.deleteTargetId) {
                deleteClient(window.deleteTargetId);
                refresh();
                showToast('Client supprimé');
                $('#confirm-modal').classList.add('hidden');
                window.deleteTargetId = null;
            }
        });

        $('#btn-parse-save').addEventListener('click', async () => {
            const text = $('#paste-area').value.trim();
            if (!text) {
                showToast('Veuillez coller du texte d\'abord', 'info');
                return;
            }

            const data = parseWhatsApp(text);
            // Ensure the date from the header is used, as WhatsApp paste doesn't provide one
            data.date = $('#add-screen-date').value || viewDate || todayStr();

            if (!data.nomClient && !data.telephone) {
                showToast('Format non reconnu. Assurez-vous d\'inclure au moins le Nom et le Téléphone.', 'error');
                return;
            }

            const added = await checkAndAddClient(data);
            if (added) {
                showToast('Client extrait et ajouté avec succès ! ✅');
                $('#paste-area').value = '';

                // Synchroniser les filtres de date
                viewDate = data.date;
                if ($('#clients-start-date')) $('#clients-start-date').value = data.date;
                if ($('#clients-end-date')) $('#clients-end-date').value = data.date;
                switchScreen('clients');
            }
        });

        // ─── Clear Paste Area ───
        const btnClearPaste = $('#btn-clear-paste');
        if (btnClearPaste) {
            btnClearPaste.addEventListener('click', () => {
                const area = $('#paste-area');
                if (area) {
                    area.value = '';
                    area.focus();
                    showToast('Zone de collage vidée', 'info');
                }
            });
        }

        // ─── Extraction Security Modal Events ───
        const btnXtSecConfirm = $('#xt-security-confirm');
        if (btnXtSecConfirm) {
            btnXtSecConfirm.addEventListener('click', validateExtractionKey);
        }
        
        const btnXtSecCancel = $('#xt-security-cancel');
        if (btnXtSecCancel) {
            btnXtSecCancel.addEventListener('click', () => {
                $('#xt-security-modal').classList.add('hidden');
                showToast('Accès refusé', 'info');
            });
        }
        
        const xtSecInput = $('#xt-security-key');
        if (xtSecInput) {
            xtSecInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') validateExtractionKey();
            });
        }

        refresh();
    }

    // ─── Global window functions ───
    window.cycleStatus = function (id) {
        const clients = getAllClients();
        const c = clients.find(cl => cl.id === id);
        if (!c) return;
        const workflow = ['en_attente', 'en_cours', 'effectue', 'annule'];
        c.statut = workflow[(workflow.indexOf(c.statut) + 1) % workflow.length];
        saveAllClients(clients);
        refresh();
    };

    window.openEditModal = function (id) {
        const c = getAllClients().find(cl => cl.id === id);
        if (!c) return;
        $('#edit-id').value = c.id;
        $('#edit-date').value = c.date; // Sauvegarder la date d'origine
        $('#edit-code-client').value = c.codeClient || '';
        $('#edit-code-agent').value = c.codeAgent || '';
        $('#edit-code-marketeur').value = c.codeMarketeur || '';
        $('#edit-nom-client').value = c.nomClient || '';
        $('#edit-telephone').value = c.telephone || '';
        $('#edit-intervention').value = c.intervention || '';
        if ($('#edit-priorite')) $('#edit-priorite').value = c.priorite;
        if ($('#edit-statut')) $('#edit-statut').value = c.statut;
        $('#edit-montant-base').value = c.montantBase;
        $('#edit-commission').value = c.commission || 0;
        $('#edit-montant-annexes').value = c.fraisAnnexes;
        $('#edit-montant-deplacement').value = c.deplacement;
        $('#edit-commentaire').value = c.commentaire || '';
        if ($('#edit-code-s')) $('#edit-code-s').value = c.codeS || '';
        $('#edit-modal').classList.remove('hidden');
    };

    window.closeEditModal = () => $('#edit-modal').classList.add('hidden');
    window.deleteTargetId = null;
    window.askDelete = (id) => {
        window.deleteTargetId = id;
        $('#confirm-modal').classList.remove('hidden');
    };
    window.viewArchive = (date) => {
        currentHistoryDate = date;
        switchScreen('history-detail');
    };

    function showToast(msg, type = 'success') {
        const t = document.createElement('div');
        t.className = `toast toast-${type}`;
        t.textContent = msg;
        $('#toast-container').appendChild(t);
        setTimeout(() => t.remove(), 3000);
    }

    function exportToExcel() {
        const start = ($('#clients-start-date') && $('#clients-start-date').value) ||
                      ($('#dash-start-date')   && $('#dash-start-date').value) || todayStr();
        const end   = ($('#clients-end-date')  && $('#clients-end-date').value) ||
                      ($('#dash-end-date')     && $('#dash-end-date').value) || todayStr();
        const clients = getAllClients().filter(c => c.date >= start && c.date <= end);

        if (clients.length === 0) {
            showToast('Aucune donnée à exporter pour cette période', 'info');
            return;
        }

        const rows = clients.map(c => ({
            'Date':          c.date,
            'Code Client':   c.codeClient || '-',
            'Code Agent':    c.codeAgent || '-',
            'Code Marketeur':c.codeMarketeur || '-',
            'Nom Client':    c.nomClient || '-',
            'Contact Client':c.telephone || '-',
            'Prestation':    c.intervention || '-',
            'Facture':       Number(c.montantBase) || 0,
            'Commission':    Number(c.commission) || 0,
            'Frais Annexes': Number(c.fraisAnnexes) || 0,
            'Déplacement':   Number(c.deplacement) || 0,
            'Commentaire':   c.commentaire || '-',
            'CODE S':        c.codeS || '-',
            'Statut':        c.statut || '-'
        }));

        const ws = XLSX.utils.json_to_sheet(rows);

        // Largeurs de colonnes adaptées
        ws['!cols'] = [
            { wch: 12 }, // Date
            { wch: 14 }, // Code Client
            { wch: 14 }, // Code Agent
            { wch: 16 }, // Code Marketeur
            { wch: 20 }, // Nom
            { wch: 16 }, // Contact
            { wch: 28 }, // Prestation
            { wch: 14 }, // Facture
            { wch: 14 }, // Commission
            { wch: 14 }, // Frais
            { wch: 14 }, // Dépl
            { wch: 25 }, // Commentaire
            { wch: 15 }, // CODE S
            { wch: 12 }  // Statut
        ];

        // Onglet Résumé
        const done      = clients.filter(c => c.statut === 'effectue').length;
        const totBase   = clients.reduce((s, c) => s + (Number(c.montantBase)  || 0), 0);
        const totCom    = clients.reduce((s, c) => s + (Number(c.commission)   || 0), 0);
        const totAnn    = clients.reduce((s, c) => s + (Number(c.fraisAnnexes) || 0), 0);
        const totDepl   = clients.reduce((s, c) => s + (Number(c.deplacement)  || 0), 0);
        const rate      = clients.length > 0 ? Math.round((done / clients.length) * 100) : 0;

        const summaryRows = [
            { 'Indicateur': 'Période', 'Valeur': `${start} au ${end}` },
            { 'Indicateur': 'Généré le', 'Valeur': todayStr() },
            { 'Indicateur': '' , 'Valeur': '' },
            { 'Indicateur': 'Total Fiches', 'Valeur': clients.length },
            { 'Indicateur': 'Réalisées', 'Valeur': done },
            { 'Indicateur': 'Taux de réalisation', 'Valeur': `${rate}%` },
            { 'Indicateur': '', 'Valeur': '' },
            { 'Indicateur': 'Total Factures', 'Valeur': totBase },
            { 'Indicateur': 'Total Commissions', 'Valeur': totCom },
            { 'Indicateur': 'Total Frais Annexes', 'Valeur': totAnn },
            { 'Indicateur': 'Total Déplacement', 'Valeur': totDepl }
        ];
        const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
        wsSummary['!cols'] = [{ wch: 22 }, { wch: 20 }];

        const wb = XLSX.utils.book_new();

        // Ajouter une ligne de totaux en bas du tableau de données
        XLSX.utils.sheet_add_aoa(ws, [[
            'TOTAUX', '', '', '', '', '', '',
            totBase, totCom, totAnn, totDepl, '', ''
        ]], { origin: -1 });

        // Style de la ligne de totaux (gras via un commentaire indicatif)
        XLSX.utils.book_append_sheet(wb, ws, 'Données');
        XLSX.utils.book_append_sheet(wb, wsSummary, 'Résumé');
        XLSX.writeFile(wb, `rapport_megapro_${start}_${end}.xlsx`);
        showToast('Export Excel réussi ✅');
    }

    function exportToPDF() {
        const start = ($('#clients-start-date') && $('#clients-start-date').value) ||
                      ($('#dash-start-date')   && $('#dash-start-date').value) || todayStr();
        const end   = ($('#clients-end-date')  && $('#clients-end-date').value) ||
                      ($('#dash-end-date')     && $('#dash-end-date').value) || todayStr();
        const clients = getAllClients().filter(c => c.date >= start && c.date <= end);

        if (clients.length === 0) {
            showToast('Aucune donnée à exporter pour cette période', 'info');
            return;
        }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('landscape', 'mm', 'a4');
        const pageW = doc.internal.pageSize.getWidth();

        // ── En-tête ──
        doc.setFillColor(30, 20, 60);
        doc.rect(0, 0, pageW, 30, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFontSize(18);
        doc.setFont('helvetica', 'bold');
        doc.text('MEGA PRO', 14, 12);
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        doc.text('Suivi Tech – Rapport Clients', 14, 19);
        doc.setFontSize(9);
        doc.text(`Période : ${start} au ${end}`, 14, 26);
        doc.text(`Généré le : ${todayStr()}`, pageW - 14, 26, { align: 'right' });

        // ── Résumé financier ──
        const done    = clients.filter(c => c.statut === 'effectue').length;
        const totBase = clients.reduce((s, c) => s + (Number(c.montantBase)  || 0), 0);
        const totCom  = clients.reduce((s, c) => s + (Number(c.commission)   || 0), 0);
        const totAnn  = clients.reduce((s, c) => s + (Number(c.fraisAnnexes) || 0), 0);
        const totDepl = clients.reduce((s, c) => s + (Number(c.deplacement)  || 0), 0);
        const rate    = clients.length > 0 ? Math.round((done / clients.length) * 100) : 0;

        doc.setTextColor(30, 20, 60);
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        const summaryY = 38;
        doc.setFillColor(245, 243, 255);
        doc.roundedRect(10, summaryY - 4, pageW - 20, 18, 2, 2, 'F');

        const summItems = [
            `${clients.length} fiches`,
            `${done} realisees (${rate}%)`,
            `Factures : ${formatMoneyPDF(totBase)} F`,
            `Commissions : ${formatMoneyPDF(totCom)} F`,
            `Frais Ann. : ${formatMoneyPDF(totAnn)} F`,
            `Depl. : ${formatMoneyPDF(totDepl)} F`
        ];
        const colWidth = (pageW - 20) / summItems.length;
        summItems.forEach((txt, i) => {
            doc.setFont('helvetica', i > 1 ? 'normal' : 'bold');
            doc.text(txt, 14 + i * colWidth, summaryY + 5);
        });

        // ── Tableau ──
        const headers = ['Date', 'Code C.', 'Code Ag.', 'Code Mk.', 'Nom', 'Contact', 'Prestation', 'Facture', 'Com.', 'Frais', 'Dépl.', 'Obs.', 'CODE S', 'Statut'];
        const rows = clients.map(c => [
            c.date || '-',
            c.codeClient || '-',
            c.codeAgent || '-',
            c.codeMarketeur || '-',
            (c.nomClient || '-').substring(0, 15),
            c.telephone  || '-',
            (c.intervention || '-').substring(0, 20),
            formatMoneyPDF(c.montantBase),
            formatMoneyPDF(c.commission),
            formatMoneyPDF(c.fraisAnnexes),
            formatMoneyPDF(c.deplacement),
            (c.commentaire || '-').substring(0, 15),
            c.codeS || '-',
            c.statut || '-'
        ]);

        doc.autoTable({
            head: [headers],
            body: rows,
            startY: summaryY + 18,
            theme: 'grid',
            headStyles: {
                fillColor:  [60, 30, 120],
                textColor:  255,
                fontStyle:  'bold',
                fontSize:   7,
                halign:     'center'
            },
            bodyStyles: { fontSize: 6.5, cellPadding: 1.5 },
            alternateRowStyles: { fillColor: [248, 245, 255] },
            columnStyles: {
                0: { cellWidth: 16 },
                1: { cellWidth: 16 },
                2: { cellWidth: 16 },
                3: { cellWidth: 16 },
                4: { cellWidth: 26 },
                5: { cellWidth: 22 },
                6: { cellWidth: 'auto' },
                7: { halign: 'right', cellWidth: 18 },
                8: { halign: 'right', cellWidth: 16 },
                9: { halign: 'right', cellWidth: 16 },
                10: { halign: 'right', cellWidth: 16 },
                11: { cellWidth: 20 },
                12: { cellWidth: 15 },
                13: { cellWidth: 15, halign: 'center' }
            },
            didDrawPage: (data) => {
                // Pied de page
                const pageCount = doc.internal.getNumberOfPages();
                doc.setFontSize(7);
                doc.setTextColor(150);
                doc.text(
                    `Page ${data.pageNumber} / ${pageCount}  –  Mega Pro © ${new Date().getFullYear()}`,
                    pageW / 2, doc.internal.pageSize.getHeight() - 6,
                    { align: 'center' }
                );
            }
        });

        doc.save(`rapport_megapro_${start}_${end}.pdf`);
        showToast('Export PDF réussi ✅');
    }

    function exportToCSV() {
        const clients = getAllClients();
        if (clients.length === 0) {
            showToast('Aucune donnée à exporter', 'info');
            return;
        }
        const headers = ['Date', 'Code Client', 'Code Agent', 'Code Marketeur', 'Nom Client', 'Contact Client', 'Prestation', 'Facture', 'Commission', 'Frais Annexes', 'Déplacement', 'Commentaire', 'CODE S', 'Statut'];
        let csvContent = '\uFEFF' + headers.join(';') + '\n';
        clients.forEach(c => {
            const row = [c.date, c.codeClient, c.codeAgent, c.codeMarketeur, c.nomClient, c.telephone, c.intervention, c.montantBase, c.commission, c.fraisAnnexes, c.deplacement, c.commentaire, c.codeS, c.statut]
                .map(v => `"${(v || '').toString().replace(/"/g, '""')}"`)
                .join(';');
            csvContent += row + '\n';
        });

        // Ligne de totaux
        const totBase = clients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const totCom  = clients.reduce((s, c) => s + (Number(c.commission) || 0), 0);
        const totAnn  = clients.reduce((s, c) => s + (Number(c.fraisAnnexes) || 0), 0);
        const totDepl = clients.reduce((s, c) => s + (Number(c.deplacement) || 0), 0);
        const totalsRow = ['"TOTAUX"', '""', '""', '""', '""', '""', '""', `"${totBase}"`, `"${totCom}"`, `"${totAnn}"`, `"${totDepl}"`, '""', '""', '""'].join(';');
        csvContent += totalsRow + '\n';
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `export_megapro_${todayStr()}.csv`;
        link.click();
        URL.revokeObjectURL(url);
        showToast('Exportation réussie');
    }

    function exportToJSON() {
        const clients = getAllClients();
        if (clients.length === 0) {
            showToast('Aucune donnée à exporter', 'info');
            return;
        }

        const dataStr = JSON.stringify(clients, null, 2);
        const blob = new Blob([dataStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `backup_megapro_${todayStr()}.json`;
        link.click();
        URL.revokeObjectURL(url);

        showToast('Sauvegarde JSON réussie ✅');
    }

    // ─── Import Functions ───

    // Column name mapping: exported header → internal field
    const HEADER_MAP = {
        'date': 'date',
        'code client': 'codeClient', 'code': 'codeClient', 'id client': 'codeClient',
        'code agent': 'codeAgent', 'code ag.': 'codeAgent',
        'code marketeur': 'codeMarketeur', 'code mk.': 'codeMarketeur', 'marketeur': 'codeMarketeur',
        'nom client': 'nomClient', 'nom': 'nomClient',
        'téléphone': 'telephone', 'telephone': 'telephone', 'tél.': 'telephone', 'tel': 'telephone', 'contact': 'telephone', 'contact client': 'telephone',
        'service': 'intervention', 'intervention': 'intervention', 'prestation': 'intervention',
        'facture': 'montantBase', 'facture base': 'montantBase', 'montant base': 'montantBase', 'facture (f)': 'montantBase',
        'facture ttc': 'montantTTC', '🟢 facture ttc': 'montantTTC',
        'commission': 'commission', 'commission (f)': 'commission', 'com.': 'commission',
        'frais annexes': 'fraisAnnexes', 'f. ann.': 'fraisAnnexes', 'annexe': 'fraisAnnexes', 'annexes': 'fraisAnnexes', 'frais ann. (f)': 'fraisAnnexes',
        'déplacement': 'deplacement', 'deplacement': 'deplacement', 'dépl.': 'deplacement', 'trajet': 'deplacement', 'dépl. (f)': 'deplacement',
        'commentaire': 'commentaire', 'obs.': 'commentaire',
        'code s': 'codeS',
        'priorité': 'priorite', 'priorite': 'priorite',
        'statut': 'statut', 'status': 'statut'
    };

    function mapRowToClient(row) {
        const client = {
            codeClient: '', codeAgent: '', codeMarketeur: '', codeS: '',
            nomClient: '', telephone: '', intervention: '',
            montantBase: 0, commission: 0, fraisAnnexes: 0, deplacement: 0,
            commentaire: '',
            priorite: 'normal', statut: 'en_attente', date: todayStr()
        };
        for (const [key, value] of Object.entries(row)) {
            const normalizedKey = key.trim().toLowerCase();
            const field = HEADER_MAP[normalizedKey];
            if (!field) continue;

            if (['montantBase', 'commission', 'fraisAnnexes', 'deplacement'].includes(field)) {
                // Parse numeric: remove currency symbols, spaces, handle French formatting
                let num = String(value || '').replace(/[^\d.,-]/g, '').replace(/\s/g, '');
                // Handle French decimal separator
                if (num.includes(',') && !num.includes('.')) {
                    num = num.replace(',', '.');
                }
                client[field] = Number(num) || 0;
            } else if (field === 'date') {
                // Try to normalize date
                const dateStr = String(value || '').trim();
                client[field] = normalizeDate(dateStr) || todayStr();
            } else {
                client[field] = String(value || '').trim();
            }
        }
        return client;
    }

    function importFromSpreadsheet(file, formatName) {
        const reader = new FileReader();
        reader.onload = async function (event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

                if (!rows || rows.length === 0) {
                    showToast('Le fichier ne contient aucune donnée.', 'error');
                    return;
                }

                let importedCount = 0;
                for (const row of rows) {
                    const clientData = mapRowToClient(row);
                    // Skip rows that have no identifying data at all
                    if (!clientData.codeClient && !clientData.telephone && !clientData.intervention) continue;
                    const added = await checkAndAddClient(clientData);
                    if (added) importedCount++;
                }

                refresh();
                showToast(`Import ${formatName} réussi : ${importedCount} client(s) ajouté(s) ✅`);
            } catch (err) {
                console.error(err);
                showToast(`Erreur lors de l'import ${formatName}.`, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function importFromPDF(file) {
        const reader = new FileReader();
        reader.onload = async function (event) {
            try {
                if (typeof pdfjsLib === 'undefined') {
                    showToast('La bibliothèque PDF.js n\'est pas disponible.', 'error');
                    return;
                }
                pdfjsLib.GlobalWorkerOptions.workerSrc = '';

                const typedArray = new Uint8Array(event.target.result);
                const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;

                let fullText = '';
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }

                const PDF_HEADERS = ['Date', 'Code', 'Tél.', 'Service', 'Facture', 'F. Ann.', 'TTC', 'Dépl.', 'Priorité', 'Statut'];
                const lines = fullText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                let importedCount = 0;

                for (const line of lines) {
                    const dateMatch = line.match(/(\d{4}-\d{2}-\d{2})/);
                    if (!dateMatch) continue;
                    if (line.includes('Mega Pro') || line.includes('Généré le')) continue;

                    const parts = line.split(/\s{2,}|\t/);
                    if (parts.length < 5) continue;

                    const row = {};
                    PDF_HEADERS.forEach((header, idx) => {
                        if (idx < parts.length) {
                            row[header] = parts[idx].trim();
                        }
                    });

                    const clientData = mapRowToClient(row);
                    if (!clientData.codeClient && !clientData.telephone && !clientData.intervention) continue;
                    const added = await checkAndAddClient(clientData);
                    if (added) importedCount++;
                }

                if (importedCount === 0) {
                    const allText = fullText.replace(/\n/g, ' ');
                    const dateSegments = allText.split(/(?=\d{4}-\d{2}-\d{2})/);

                    for (const segment of dateSegments) {
                        const dm = segment.match(/^(\d{4}-\d{2}-\d{2})/);
                        if (!dm) continue;

                        const rest = segment.slice(10).trim();
                        const tokens = rest.split(/\s+/);
                        if (tokens.length < 4) continue;

                        const clientData = {
                            date: dm[1],
                            codeClient: tokens[0] || '',
                            telephone: tokens[1] || '',
                            intervention: tokens[2] || '',
                            montantBase: Number(String(tokens[3] || '0').replace(/[^\d]/g, '')) || 0,
                            fraisAnnexes: Number(String(tokens[4] || '0').replace(/[^\d]/g, '')) || 0,
                            commission: Number(String(tokens[5] || '0').replace(/[^\d]/g, '')) || 0,
                            deplacement: Number(String(tokens[6] || '0').replace(/[^\d]/g, '')) || 0,
                            priorite: tokens[7] || 'normal',
                            statut: tokens[8] || 'en_attente'
                        };

                        if (!clientData.codeClient && !clientData.telephone) continue;
                        addClient(clientData);
                        importedCount++;
                    }
                }

                refresh();
                if (importedCount > 0) {
                    showToast(`Import PDF réussi : ${importedCount} client(s) ajouté(s) ✅`);
                } else {
                    showToast('Aucune donnée client trouvée dans le PDF.', 'error');
                }
            } catch (err) {
                console.error(err);
                showToast('Erreur lors de l\'import PDF.', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function importExtractionFromJSON(file) {
        const reader = new FileReader();
        reader.onload = function (event) {
            try {
                const importedData = JSON.parse(event.target.result);
                if (!Array.isArray(importedData)) {
                    throw new Error('Format invalide : le fichier doit contenir un tableau de clients.');
                }

                const newRows = importedData.map(c => ({
                    id: generateId(),
                    codeClient: c.codeClient || '',
                    codeAgent: c.codeAgent || '',
                    codeMarketeur: c.codeMarketeur || '',
                    nomClient: c.nomClient || c.nom || '',
                    contact: c.contact || c.telephone || '',
                    facture: Number(c.facture || c.montantBase) || 0,
                    commission: Number(c.commission) || 0,
                    fraisAnnexes: Number(c.fraisAnnexes) || 0,
                    deplacement: Number(c.deplacement) || 0,
                    commentaire: c.commentaire || '',
                    date: c.date || todayStr()
                }));

                extractedRows = [...extractedRows, ...newRows];
                renderExtractionTable();
                showToast(`Import JSON réussi : ${newRows.length} fiche(s) ajoutée(s) ✅`);
            } catch (err) {
                console.error(err);
                showToast('Erreur lors de l\'import JSON : ' + err.message, 'error');
            }
        };
        reader.readAsText(file);
    }

    /** Importe un fichier Excel (.xlsx) dans l'Outil d'Extraction */
    function importExtractionFromExcel(file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const wb   = XLSX.read(e.target.result, { type: 'array' });
                const ws   = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
                if (!rows.length) { showToast('Fichier Excel vide ou non reconnu', 'info'); return; }

                const newRows = rows.map(r => ({
                    id:           generateId(),
                    codeClient:   r['Code Client'] || r['code client'] || r['Code C.'] || '',
                    codeAgent:    r['Code Agent'] || r['Code Ag.'] || '',
                    codeMarketeur:r['Code Marketeur'] || r['Code Mk.'] || '',
                    nomClient:    r['Nom Client'] || r['Nom'] || '',
                    contact:      r['Contact Client'] || r['Contact'] || r['Téléphone'] || '',
                    service:      r['Prestation'] || r['Service'] || r['Intervention'] || '',
                    facture:      parseAmt(r['Facture (F)'] || r['Facture'] || 0),
                    commission:   parseAmt(r['Commission (F)'] || r['Commission'] || r['Com.'] || 0),
                    fraisAnnexes: parseAmt(r['Frais Ann. (F)'] || r['Frais Annexes'] || r['Frais'] || 0),
                    deplacement:  parseAmt(r['Dépl. (F)'] || r['Déplacement (F)'] || r['Déplacement'] || r['Dépl.'] || 0),
                    commentaire:  r['Commentaire'] || r['Obs.'] || '',
                    date:         r['Date'] || todayStr()
                }));

                extractedRows = [...extractedRows, ...newRows];
                renderExtractionTable();
                showToast(`Import Excel réussi : ${newRows.length} fiche(s) ajoutée(s) ✅`);
            } catch (err) {
                console.error(err);
                showToast('Erreur lors de l\'import Excel : ' + err.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    /** Importe un fichier CSV dans l'Outil d'Extraction */
    function importExtractionFromCSV(file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const wb   = XLSX.read(e.target.result, { type: 'string', raw: true });
                const ws   = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
                if (!rows.length) { showToast('Fichier CSV vide ou non reconnu', 'info'); return; }

                const newRows = rows.map(r => ({
                    id:           generateId(),
                    codeClient:   r['Code Client'] || r['code client'] || r['Code C.'] || '',
                    codeAgent:    r['Code Agent'] || r['Code Ag.'] || '',
                    codeMarketeur:r['Code Marketeur'] || r['Code Mk.'] || '',
                    nomClient:    r['Nom Client'] || r['Nom'] || '',
                    contact:      r['Contact Client'] || r['Contact'] || r['Téléphone'] || '',
                    service:      r['Prestation'] || r['Service'] || r['Intervention'] || '',
                    facture:      parseAmt(r['Facture'] || r['Facture Base'] || r['Facture (F)'] || 0),
                    commission:   parseAmt(r['Commission (F)'] || r['Commission'] || r['Com.'] || 0),
                    fraisAnnexes: parseAmt(r['Frais Annexes'] || r['Frais Ann. (F)'] || r['Frais'] || 0),
                    deplacement:  parseAmt(r['Déplacement'] || r['Dépl. (F)'] || r['Déplacement (F)'] || r['Dépl.'] || 0),
                    commentaire:  r['Commentaire'] || r['Obs.'] || '',
                    date:         r['Date'] || todayStr()
                }));

                extractedRows = [...extractedRows, ...newRows];
                renderExtractionTable();
                showToast(`Import CSV réussi : ${newRows.length} fiche(s) ajoutée(s) ✅`);
            } catch (err) {
                console.error(err);
                showToast('Erreur lors de l\'import CSV : ' + err.message, 'error');
            }
        };
        reader.readAsText(file, 'utf-8');
    }

    /** Exporte les données d'extraction actuelles en fichier JSON */
    function exportExtractionJSON() {
        if (!extractedRows || extractedRows.length === 0) {
            showToast('Aucune donnée à exporter', 'info');
            return;
        }
        const dataStr = JSON.stringify(extractedRows, null, 2);
        const blob    = new Blob([dataStr], { type: 'application/json' });
        const url     = URL.createObjectURL(blob);
        const link    = document.createElement('a');
        link.href     = url;
        link.download = `base_extraction_${todayStr()}.json`;
        link.click();
        URL.revokeObjectURL(url);
        showToast('Base de données exportée en JSON ✅');
    }

    function exportPerformancePDF() {
        const start = $('#perf-start-date').value;
        const end = $('#perf-end-date').value;
        if (!start || !end) {
            showToast('Sélectionnez d\'abord une période', 'error');
            return;
        }

        const all = getAllClients().filter(c => c.date >= start && c.date <= end);
        if (all.length === 0) {
            showToast('Aucune donnée sur cette période', 'info');
            return;
        }

        const done = all.filter(c => c.statut === 'effectue');
        const lossClients = all.filter(c => c.statut !== 'effectue' && c.statut !== 'annule');

        const totalExpected = all.reduce((s, c) => s + calcTTC(c), 0);
        const totalRealised = done.reduce((s, c) => s + calcTTC(c), 0);
        const rate = totalExpected > 0 ? Math.round((totalRealised / totalExpected) * 100) : 0;
        const lossAmount = lossClients.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const score = Math.round((rate / 100) * 20);

        const totalFactures = all.reduce((s, c) => s + (Number(c.montantBase) || 0), 0);
        const totalAnnexes = all.reduce((s, c) => s + (Number(c.fraisAnnexes) || 0), 0);
        const totalDeplacement = all.reduce((s, c) => s + (Number(c.deplacement) || 0), 0);
        const totalTTC = totalExpected;

        const counts = {
            en_attente: all.filter(c => c.statut === 'en_attente').length,
            en_cours: all.filter(c => c.statut === 'en_cours').length,
            effectue: done.length,
            annule: all.filter(c => c.statut === 'annule').length
        };

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('portrait', 'mm', 'a4');

        let y = 20;

        // Title
        doc.setFontSize(22);
        doc.setTextColor(0, 145, 255); // Brand color
        doc.text('Bilan de Performance Technicien', 14, y);
        y += 8;

        // Period
        doc.setFontSize(11);
        doc.setTextColor(100, 100, 100);
        doc.text(`Période : du ${formatDateFR(start)} au ${formatDateFR(end)}`, 14, y);
        y += 12;

        // Separator
        doc.setDrawColor(220, 220, 220);
        doc.line(14, y, 196, y);
        y += 10;

        // --- SECTION 1: SYNTHESE ---
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 145, 255);
        doc.text('1. Synthèse Globale', 14, y);
        doc.setFont('helvetica', 'normal');
        y += 10;

        doc.setFontSize(13);
        doc.setTextColor(60, 60, 60);
        doc.text(`- Total Fiches (Période) : ${all.length}`, 20, y);
        y += 8;
        doc.text(`- Interventions Effectuées : ${done.length}`, 20, y);
        y += 8;
        doc.text(`- Total Factures : ${formatMoneyPDF(totalFactures)}`, 20, y);
        y += 8;
        doc.text(`- Total Frais Annexes : ${formatMoneyPDF(totalAnnexes)}`, 20, y);
        y += 8;
        doc.text(`- Total TTC : ${formatMoneyPDF(totalTTC)}`, 20, y);
        y += 8;
        doc.text(`- Total Frais de Déplacement : ${formatMoneyPDF(totalDeplacement)}`, 20, y);
        y += 8;
        doc.text(`- Chiffre d'Affaires Theorique : ${formatMoneyPDF(totalExpected)}`, 20, y);
        y += 8;
        doc.text(`- Chiffre d'Affaires Realise : ${formatMoneyPDF(totalRealised)}`, 20, y);
        y += 10;
        
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(147, 51, 234); // purple highlight
        doc.text(`- Taux de réalisation : ${rate}%`, 20, y);
        doc.setFont('helvetica', 'normal');
        y += 12;

        // Score Highlight
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.setFillColor(243, 232, 255); // light purple background
        doc.rect(14, y, 90, 14, 'F');
        doc.setTextColor(107, 33, 168); // dark purple text
        doc.text(`Note de Performance : ${score} / 20`, 18, y + 10);
        doc.setFont('helvetica', 'normal');
        y += 24;

        // --- SECTION 2: RÉPARTITION DES STATUTS ---
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 145, 255);
        doc.text('2. Répartition par Statut', 14, y);
        doc.setFont('helvetica', 'normal');
        y += 10;

        const statusData = [
            ['Statut', 'Quantité', 'Pourcentage'],
            ['En attente', counts.en_attente, Math.round((counts.en_attente / all.length) * 100) + '%'],
            ['En cours', counts.en_cours, Math.round((counts.en_cours / all.length) * 100) + '%'],
            ['Effectué', counts.effectue, Math.round((counts.effectue / all.length) * 100) + '%'],
            ['Annulé', counts.annule, Math.round((counts.annule / all.length) * 100) + '%']
        ];

        doc.autoTable({
            startY: y,
            head: [statusData[0]],
            body: statusData.slice(1),
            theme: 'striped',
            headStyles: { fillColor: [0, 145, 255], fontSize: 11 },
            bodyStyles: { fontSize: 11 },
            margin: { left: 14 }
        });

        y = doc.lastAutoTable.finalY + 20;

        // --- SECTION 3: DÉPERDITION ---
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 145, 255);
        doc.text('3. Déperdition', 14, y);
        doc.setFont('helvetica', 'normal');
        y += 10;

        doc.setFontSize(13);
        doc.setTextColor(220, 38, 38); // Red text
        doc.text(`- Fiches non traitées : ${lossClients.length}`, 20, y);
        y += 8;
        doc.setFont('helvetica', 'bold');
        doc.text(`- Montant potentiel perdu : ${formatMoneyPDF(lossAmount)}`, 20, y);
        doc.setFont('helvetica', 'normal');

        // Footer
        doc.setFontSize(9);
        doc.setTextColor(150, 150, 150);
        doc.text(`Ce bilan a été généré automatiquement par Mega Pro Suivi Tech V1 le ${new Date().toLocaleString('fr-FR')}.`, 14, 280);

        doc.save(`bilan_performance_${todayStr()}.pdf`);
        showToast('Bilan PDF généré ✅');
    }


    // =====================================================
    //  OUTIL D'EXTRACTION — Module autonome
    // =====================================================

    // ─── Extraction State (session-only, not persisted) ───
    let extractedRows = [];

    // ─── Parsing helpers ───

    /**
     * Remove emojis and decoration from a string, return clean text
     */
    function stripDecorations(str) {
        return str
            // Remove emoji ranges
            .replace(/[\u{1F000}-\u{1FFFF}]/gu, '')
            .replace(/[\u{2600}-\u{26FF}]/gu, '')
            .replace(/[\u{2700}-\u{27BF}]/gu, '')
            .replace(/[\u{FE00}-\u{FE0F}]/gu, '')
            // Remove common box-drawing / separator characters (kept em-dash out of code)
            .replace(/[—–─═●•◦▸▪◆◇■□▲△▼▽★☆♐♀♂❤]/g, ' ')
            // Collapse multiple spaces
            .replace(/\s+/g, ' ')
            .trim();
    }

    /**
     * Normalise a monetary string to integer:
     * "15 000 FCFA" → 15000, "FCFA" → 0, "" → 0
     */
    function parseMoney(str) {
        if (!str) return 0;
        // Remove FCFA and all non-digit chars (except leading minus)
        const cleaned = str.replace(/fcfa/gi, '').replace(/[^0-9]/g, '');
        return cleaned ? parseInt(cleaned, 10) : 0;
    }

    /**
     * Try to extract the value after a "label : value" pattern on a given line.
     * Returns null if the line doesn't match the labels.
     */
    function extractField(line, labels) {
        const lc = stripDecorations(line).toLowerCase();
        for (const label of labels) {
            if (lc.includes(label)) {
                // Find colon separator (after the label)
                const colonIdx = line.indexOf(':');
                if (colonIdx !== -1) {
                    return line.slice(colonIdx + 1).trim();
                }
            }
        }
        return null;
    }

    /**
     * Parse a single fiche text and return a structured object.
     */
    function parseClientFiche(text) {
        const lines = text.split('\n');
        const row = {
            id: generateId(),
            date: '',
            codeClient: '',
            codeAgent: '',
            codeMarketeur: '',
            nomClient: '',
            telephone: '',
            intervention: '',
            montantBase: 0,
            commission: 0,
            fraisAnnexes: 0,
            deplacement: 0,
            commentaire: '',
            codeS: ''
        };

        lines.forEach(line => {
            if (!line.trim()) return;

            let val;

            val = extractField(line, ['code client', 'id client', 'code_client']);
            if (val !== null) { row.codeClient = stripDecorations(val).replace(/:/g, '').trim(); return; }

            val = extractField(line, ['code agent', 'agent']);
            if (val !== null && !line.toLowerCase().includes('marketeur')) { row.codeAgent = stripDecorations(val).replace(/:/g, '').trim(); return; }

            val = extractField(line, ['code marketeur', 'marketeur', 'mk']);
            if (val !== null) { row.codeMarketeur = stripDecorations(val).replace(/:/g, '').trim(); return; }

            val = extractField(line, ['nom client', 'nom_client', 'client']);
            if (val !== null && !line.toLowerCase().includes('code')) { row.nomClient = stripDecorations(val); return; }

            val = extractField(line, ['contact', 'tél', 'tel', 'téléphone', 'telephone']);
            if (val !== null) { row.telephone = stripDecorations(val); return; }

            val = extractField(line, ['date']);
            if (val !== null) {
                const cleaned = stripDecorations(val).split(/\s/)[0];
                const d = normalizeDate(cleaned);
                if (d) { row.date = d; return; }
            }

            val = extractField(line, ['service', 'intervention', 'prestation', 'prest']);
            if (val !== null) { row.intervention = stripDecorations(val); return; }

            val = extractField(line, ['frais de déplacement', 'frais deplacement', 'déplacement', 'deplacement', 'transport']);
            if (val !== null) { row.deplacement = parseMoney(val); return; }

            val = extractField(line, ['frais annexes', 'annexes', 'frais ann']);
            if (val !== null) { row.fraisAnnexes = parseMoney(val); return; }

            val = extractField(line, ['commission', 'com']);
            if (val !== null) { row.commission = parseMoney(val); return; }

            val = extractField(line, ['facture', 'prix']);
            if (val !== null && !line.toLowerCase().includes('ttc') && !line.toLowerCase().includes('annexes')) {
                row.montantBase = parseMoney(val); return;
            }

            val = extractField(line, ['commentaire', 'obs', 'observation', 'remarque']);
            if (val !== null) { row.commentaire = stripDecorations(val); return; }

            val = extractField(line, ['code s', 'codes', 'code_s']);
            if (val !== null) { row.codeS = stripDecorations(val); return; }
        });

        return row;
    }

    function parseMultipleFiches(text) {
        // Splitting into individual fiches
        // Safest approach: split by double blank lines OR "Fiche Client" / "🔴" patterns
        // We will process the text by looking for clear start boundaries.
        
        let rawFiches = text
            // Split whenever we see exactly the "🔴—📫Fiche Client" or similar distinct header
            .split(/(?=🔴[—\-]*📫\s*Fiche|Fiche\s+Client(?:\s+P\(\d+\/\d+\))?|🔐\s*Code\s+Client)/i)
            .filter(s => s.trim().length > 0);

        // Fallback: if split produced only 1 part, and it's quite large, try splitting on double blank lines but carefully
        if (rawFiches.length <= 1) {
            const parts = text.split(/\n{2,}/);
            const fiches = [];
            let current = [];
            
            // Fiche usually starts with Code Client, Fiche Client, Nom, or the red circle
            const ficheStartPattern = /code\s*client|fiche\s*client|🔴/i;
            
            parts.forEach(part => {
                if (ficheStartPattern.test(part) && current.length > 0) {
                    // Check if current block already has enough data to be considered a fiche
                    const block = current.join('\n\n');
                    if (block.toLowerCase().includes('code client') || block.toLowerCase().includes('facture')) {
                        fiches.push(block);
                        current = [];
                    }
                }
                current.push(part);
            });
            if (current.length > 0) fiches.push(current.join('\n\n'));
            
            rawFiches = fiches;
        }

        return rawFiches.filter(f => f.trim()).map(parseClientFiche);
    }



    // ─── Apply filters and get visible rows ───
    function getFilteredExtractionRows() {
        const query = ($('#xt-search') || {}).value?.toLowerCase().trim() || '';
        const startDate = ($('#xt-filter-start') || {}).value || '';
        const endDate = ($('#xt-filter-end') || {}).value || '';

        return extractedRows.filter(r => {
            // Text search
            if (query) {
                const haystack = [r.codeClient, r.codeAgent, r.codeMarketeur, r.nomClient, r.telephone, r.intervention, r.commentaire, r.codeS]
                    .join(' ').toLowerCase();
                if (!haystack.includes(query)) return false;
            }
            // Date range
            if (startDate && r.date && r.date < startDate) return false;
            if (endDate && r.date && r.date > endDate) return false;
            return true;
        });
    }

    // ─── Render the extraction table ───
    function renderExtractionTable() {
        const tbody = $('#xt-tbody');
        const emptyState = $('#xt-empty-state');
        const rowCountEl = $('#xt-row-count');
        if (!tbody) return;

        const filtered = getFilteredExtractionRows();

        // Update totals
        const totFact = filtered.reduce((s, r) => s + r.montantBase, 0);
        const totAnn = filtered.reduce((s, r) => s + r.fraisAnnexes, 0);
        const totDepl = filtered.reduce((s, r) => s + r.deplacement, 0);
        const totTTC = filtered.reduce((s, r) => s + calcTTC(r), 0);

        const setEl = (id, val) => { const el = $(id); if (el) el.textContent = formatMoney(val); };
        setEl('#xt-total-facture', totFact);
        setEl('#xt-total-annexes', totAnn);
        setEl('#xt-total-deplacement', totDepl);
        setEl('#xt-total-ttc', totTTC);

        if (rowCountEl) {
            rowCountEl.textContent = `${filtered.length} fiche(s) — ${extractedRows.length} au total`;
        }

        if (filtered.length === 0) {
            tbody.innerHTML = '';
            if (emptyState) emptyState.style.display = '';
            return;
        }
        if (emptyState) emptyState.style.display = 'none';

        tbody.innerHTML = filtered.map(r => `
            <tr data-xt-id="${escapeHTML(r.id)}" class="xt-row">
                <td class="xt-cell-date">${escapeHTML(r.date || todayStr())}</td>
                <td class="xt-cell"><span class="code-pill">${escapeHTML(r.codeClient || '—')}</span></td>
                <td class="xt-cell"><span class="code-pill">${escapeHTML(r.codeAgent || '—')}</span></td>
                <td class="xt-cell-muted">${escapeHTML(r.codeMarketeur || '—')}</td>
                <td class="xt-cell-bold">${escapeHTML(r.nomClient || '—')}</td>
                <td class="xt-cell-muted">${escapeHTML(r.telephone || '—')}</td>
                <td class="xt-cell-truncate" title="${escapeHTML(r.intervention || '')}">${escapeHTML(r.intervention || '—')}</td>
                <td class="xt-cell-money">${r.montantBase > 0 ? formatMoney(r.montantBase) : '0'}</td>
                <td class="xt-cell-money-warn">${r.commission > 0 ? formatMoney(r.commission) : '0'}</td>
                <td class="xt-cell-money-muted">${r.fraisAnnexes > 0 ? formatMoney(r.fraisAnnexes) : '0'}</td>
                <td class="xt-cell-money-muted">${r.deplacement > 0 ? formatMoney(r.deplacement) : '0'}</td>
                <td class="xt-cell-truncate" style="max-width: 150px;" title="${escapeHTML(r.commentaire || '')}">${escapeHTML(r.commentaire || '—')}</td>
                <td class="xt-cell-center">
                    <button class="xt-btn-delete" onclick="window.xtDeleteRow('${escapeHTML(r.id)}')" title="Supprimer">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/></svg>
                    </button>
                </td>
            </tr>
        `).join('');
    }

    // ─── Delete a row ───
    window.xtDeleteRow = function(id) {
        extractedRows = extractedRows.filter(r => r.id !== id);
        renderExtractionTable();
    };

    // ─── Export functions ───
    function exportExtractionCSV() {
        const rows = getFilteredExtractionRows();
        if (rows.length === 0) { showToast('Aucune donnée à exporter', 'info'); return; }

        const headers = ['Date', 'Code Client', 'Code Agent', 'Code Marketeur', 'Nom Client', 'Contact Client', 'Prestation', 'Facture', 'Commission', 'Frais Annexes', 'Déplacement', 'Commentaire', 'CODE S'];
        let csv = '\uFEFF' + headers.join(';') + '\n';
        rows.forEach(r => {
            const row = [r.date, r.codeClient, r.codeAgent, r.codeMarketeur, r.nomClient, r.telephone, r.intervention,
                r.montantBase, r.commission, r.fraisAnnexes, r.deplacement, r.commentaire, r.codeS]
                .map(v => `"${(v ?? '').toString().replace(/"/g, '""')}"`).join(';');
            csv += row + '\n';
        });
        
        // Add totals row
        const totFact = rows.reduce((s, r) => s + (Number(r.montantBase) || 0), 0);
        const totCom = rows.reduce((s, r) => s + (Number(r.commission) || 0), 0);
        const totAnn = rows.reduce((s, r) => s + (Number(r.fraisAnnexes) || 0), 0);
        const totDepl = rows.reduce((s, r) => s + (Number(r.deplacement) || 0), 0);
        const totalsRow = ['"TOTAUX"', '""', '""', '""', '""', '""', '""', `"${totFact}"`, `"${totCom}"`, `"${totAnn}"`, `"${totDepl}"`, '""', '""'].join(';');
        csv += totalsRow + '\n';
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = `extraction_megapro_${todayStr()}.csv`;
        a.click(); URL.revokeObjectURL(url);
        showToast('Export CSV réussi ✅');
    }

    function exportExtractionExcel() {
        const rows = getFilteredExtractionRows();
        if (rows.length === 0) { showToast('Aucune donnée à exporter', 'info'); return; }

        const data = rows.map(r => ({
            'Date': r.date,
            'Code Client': r.codeClient,
            'Code Agent': r.codeAgent,
            'Code Marketeur': r.codeMarketeur,
            'Nom Client': r.nomClient,
            'Contact Client': r.telephone,
            'Prestation': r.intervention,
            'Facture': r.montantBase,
            'Commission': r.commission,
            'Frais Annexes': r.fraisAnnexes,
            'Frais Déplacement': r.deplacement,
            'Commentaire': r.commentaire,
            'CODE S': r.codeS
        }));

        // Add totals row
        const totals = {
            'Date': 'TOTAL', 'Code Client': '', 'Code Agent': '', 'Code Marketeur': '', 'Nom Client': '', 'Contact Client': '', 'Prestation': '',
            'Facture': rows.reduce((s, r) => s + (Number(r.montantBase) || 0), 0),
            'Commission': rows.reduce((s, r) => s + (Number(r.commission) || 0), 0),
            'Frais Annexes': rows.reduce((s, r) => s + (Number(r.fraisAnnexes) || 0), 0),
            'Frais Déplacement': rows.reduce((s, r) => s + (Number(r.deplacement) || 0), 0),
            'Commentaire': '', 'CODE S': ''
        };
        data.push(totals);

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Extraction Mega Pro');
        XLSX.writeFile(wb, `extraction_megapro_${todayStr()}.xlsx`);
        showToast('Export Excel réussi ✅');
    }

    function exportExtractionPDF() {
        const rows = getFilteredExtractionRows();
        if (rows.length === 0) { showToast('Aucune donnée à exporter', 'info'); return; }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('landscape', 'mm', 'a4');

        doc.setFontSize(18);
        doc.setTextColor(0, 145, 255);
        doc.text('Mega Pro – Extraction de Fiches Clients', 14, 16);

        doc.setFontSize(9);
        doc.setTextColor(120, 120, 120);
        doc.text(`Généré le ${new Date().toLocaleString('fr-FR')} — ${rows.length} fiche(s)`, 14, 23);

        // Totals
        const totFact = rows.reduce((s, r) => s + (Number(r.montantBase) || 0), 0);
        const totAnn = rows.reduce((s, r) => s + (Number(r.fraisAnnexes) || 0), 0);
        const totDepl = rows.reduce((s, r) => s + (Number(r.deplacement) || 0), 0);
        const totTTC = rows.reduce((s, r) => s + calcTTC(r), 0);

        doc.setFontSize(8);
        doc.setTextColor(40, 40, 40);
        doc.text(`Total Facture: ${formatMoneyPDF(totFact)}   |   Total Frais Annexes: ${formatMoneyPDF(totAnn)}   |   Total Deplacement: ${formatMoneyPDF(totDepl)}   |   Total TTC: ${formatMoneyPDF(totTTC)}`, 14, 30);

        const headers = ['Date', 'Code', 'Nom Client', 'Contact', 'Prestation', 'Facture', 'Com.', 'F. Ann.', 'Depl.', 'Obs.', 'CODE S'];
        const body = rows.map(r => [
            r.date || '', r.codeClient || '', r.nomClient || '', r.telephone || '',
            r.intervention || '',
            r.montantBase ? formatMoneyPDF(r.montantBase) : '0',
            r.commission ? formatMoneyPDF(r.commission) : '0',
            r.fraisAnnexes ? formatMoneyPDF(r.fraisAnnexes) : '0',
            r.deplacement ? formatMoneyPDF(r.deplacement) : '0',
            r.commentaire || '', r.codeS || ''
        ]);

        doc.autoTable({
            head: [headers],
            body: body,
            startY: 36,
            theme: 'grid',
            headStyles: { fillColor: [0, 145, 255], fontSize: 7, textColor: 255 },
            bodyStyles: { fontSize: 6.5 },
            alternateRowStyles: { fillColor: [240, 246, 255] },
            columnStyles: {
                6: { halign: 'right' }, 7: { halign: 'right' },
                8: { halign: 'right', fontStyle: 'bold' }, 9: { halign: 'right' }
            }
        });

        // Footer totals
        const finalY = doc.lastAutoTable.finalY + 8;
        doc.setFontSize(8);
        doc.setTextColor(0, 145, 255);
        doc.text(`Totaux : Facture=${formatMoney(totFact)} | Frais Ann.=${formatMoney(totAnn)} | Dépl.=${formatMoney(totDepl)} | TTC=${formatMoney(totTTC)}`, 14, finalY);

        doc.save(`extraction_megapro_${todayStr()}.pdf`);
        showToast('Export PDF réussi ✅');
    }

    // ─── Expose export functions globally (inline onclick in HTML) ───
    window.exportExtractionCSV = exportExtractionCSV;
    window.exportExtractionExcel = exportExtractionExcel;
    window.exportExtractionPDF = exportExtractionPDF;

    let pendingExtractions = [];


    // ─── Init extraction event listeners ───
    function initExtractionModule() {
        const modal = $('#global-duplicate-modal');
        const btnForce = $('#global-dup-force');
        const btnCancel = $('#global-dup-cancel');
        const pasteArea = $('#xt-paste-area');
        const statusEl = $('#xt-parse-status');

        function addExtractedRows(rows, countValid) {
            extractedRows = [...extractedRows, ...rows];
            renderExtractionTable();
            if (statusEl) statusEl.textContent = `✅ ${countValid} fiche(s) ajoutée(s)`;
            if (pasteArea) pasteArea.value = '';
            showToast(`${countValid} fiche(s) ajoutée(s) avec succès ✅`, 'success');
        }

        const btnParse = $('#btn-xt-parse');
        if (btnParse) {
            btnParse.addEventListener('click', () => {
                const text = pasteArea?.value?.trim() || '';
                if (!text) { showToast('Veuillez coller du texte d\'abord', 'info'); return; }

                const newRows = parseMultipleFiches(text);
                const valid = newRows.filter(r => r.codeClient || r.telephone || r.intervention);

                if (valid.length === 0) {
                    showToast('Aucune fiche reconnue. Vérifiez le format du texte.', 'error');
                    return;
                }

                const duplicates = valid.filter(newR => {
                    return extractedRows.some(oldR =>
                        (newR.codeClient && oldR.codeClient && newR.codeClient.toLowerCase() === oldR.codeClient.toLowerCase()) ||
                        (newR.telephone && oldR.telephone && newR.telephone.replace(/\s+/g, '') === oldR.telephone.replace(/\s+/g, ''))
                    );
                });

                if (duplicates.length > 0) {
                    pendingExtractions = valid;
                    if (modal && btnForce && btnCancel) {
                        const msg = $('#global-duplicate-msg');
                        if (msg) msg.textContent = `${duplicates.length} fiche(s) semble(nt) déjà exister dans l'extraction. Voulez-vous quand même les ajouter ?`;
                        modal.classList.remove('hidden');

                        const handleForce = () => {
                            cleanup();
                            addExtractedRows(pendingExtractions, pendingExtractions.length);
                            pendingExtractions = [];
                        };

                        const handleCancel = () => {
                            cleanup();
                            pendingExtractions = [];
                            showToast('Ajout annulé', 'info');
                        };

                        const cleanup = () => {
                            modal.classList.add('hidden');
                            btnForce.removeEventListener('click', handleForce);
                            btnCancel.removeEventListener('click', handleCancel);
                        };

                        btnForce.addEventListener('click', handleForce);
                        btnCancel.addEventListener('click', handleCancel);
                    }
                } else {
                    addExtractedRows(valid, valid.length);
                }
            });
        }

        // Vider toutes les fiches
        const btnClearAll = $('#btn-xt-clear-all');
        if (btnClearAll) {
            btnClearAll.addEventListener('click', () => {
                if (extractedRows.length === 0) { showToast('Rien à vider', 'info'); return; }
                extractedRows = [];
                renderExtractionTable();
                if (statusEl) statusEl.textContent = '';
                showToast('Tableau vidé', 'info');
            });
        }

        // Vider le champ de texte
        const btnClearText = $('#btn-xt-clear-text');
        if (btnClearText && pasteArea) {
            btnClearText.addEventListener('click', () => {
                pasteArea.value = '';
                if (statusEl) statusEl.textContent = '';
                pasteArea.focus();
                showToast('Zone de texte vidée', 'info');
            });
        }

        // ─── Export JSON (Outil d'Extraction) ───
        const btnXtExportJson = $('#btn-xt-export-json');
        if (btnXtExportJson) {
            btnXtExportJson.addEventListener('click', () => {
                const menu = $('#xt-export-menu');
                if (menu) menu.classList.add('hidden');
                exportExtractionJSON();
            });
        }

        // ─── Import dropdown (Outil d'Extraction) ───
        const btnImportToggle = $('#btn-xt-import-toggle');
        const xtImportMenu    = $('#xt-import-menu');

        if (btnImportToggle && xtImportMenu) {
            btnImportToggle.addEventListener('click', (e) => {
                e.stopPropagation();
                xtImportMenu.classList.toggle('hidden');
                const xtExportMenu = $('#xt-export-menu');
                if (xtExportMenu) xtExportMenu.classList.add('hidden');
            });
        }

        // Helper : brancher bouton → input file → callback
        function bindImportBtn(btnId, inputId, handler) {
            const btn   = $(`#${btnId}`);
            const input = $(`#${inputId}`);
            if (!btn || !input) return;
            btn.addEventListener('click', () => {
                if (xtImportMenu) xtImportMenu.classList.add('hidden');
                input.click();
            });
            input.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) { handler(file); e.target.value = ''; }
            });
        }

        bindImportBtn('btn-xt-import-json',  'xt-import-json-input',  importExtractionFromJSON);
        bindImportBtn('btn-xt-import-excel', 'xt-import-excel-input', importExtractionFromExcel);
        bindImportBtn('btn-xt-import-csv',   'xt-import-csv-input',   importExtractionFromCSV);
        bindImportBtn('btn-xt-import-pdf',   'xt-import-pdf-input',   (file) => {
            showToast('Import PDF non pris en charge nativement — utilisez JSON, Excel ou CSV.', 'info');
        });

        // Filter events
        ['xt-search', 'xt-filter-start', 'xt-filter-end'].forEach(id => {
            const el = $(`#${id}`);
            if (el) el.addEventListener('input', renderExtractionTable);
            if (el) el.addEventListener('change', renderExtractionTable);
        });

        // Reset filters
        const btnReset = $('#btn-xt-reset-filters');
        if (btnReset) {
            btnReset.addEventListener('click', () => {
                const ids = ['xt-search', 'xt-filter-start', 'xt-filter-end'];
                ids.forEach(id => { const el = $(`#${id}`); if (el) el.value = ''; });
                renderExtractionTable();
                showToast('Filtres réinitialisés', 'info');
            });
        }

        // Export Dropdown Toggle
        const btnExportMain = $('#btn-xt-export-main');
        const exportMenu = $('#xt-export-menu');
        const importMenu = $('#xt-import-menu');

        // Ferme tous les dropdowns de l'Outil d'Extraction
        function closeXtMenus() {
            if (exportMenu) exportMenu.classList.add('hidden');
            if (importMenu) importMenu.classList.add('hidden');
        }

        if (btnExportMain && exportMenu) {
            btnExportMain.addEventListener('click', (e) => {
                e.stopPropagation();
                const wasHidden = exportMenu.classList.contains('hidden');
                closeXtMenus();
                if (wasHidden) exportMenu.classList.remove('hidden');
            });
            exportMenu.querySelectorAll('.export-option').forEach(opt => {
                opt.addEventListener('click', () => closeXtMenus());
            });
        }

        // Fermeture du menu Import au clic sur option (le toggle est déjà géré plus haut)
        if (importMenu) {
            importMenu.querySelectorAll('.export-option').forEach(opt => {
                opt.addEventListener('click', () => closeXtMenus());
            });
        }

        // Fermeture globale au clic en dehors des deux menus
        document.addEventListener('click', (e) => {
            const xtExportWrapper = $('#xt-export-wrapper');
            const xtImportWrapper = $('#xt-import-wrapper');
            const insideExport = xtExportWrapper && xtExportWrapper.contains(e.target);
            const insideImport = xtImportWrapper && xtImportWrapper.contains(e.target);
            if (!insideExport && !insideImport) {
                closeXtMenus();
            }
        });


        // Focus styling for textarea
        if (pasteArea) {
            pasteArea.addEventListener('focus', () => pasteArea.style.borderColor = 'var(--accent)');
            pasteArea.addEventListener('blur', () => pasteArea.style.borderColor = 'var(--glass-border)');
        }

        // Initial render (empty state)
        renderExtractionTable();
    }

    document.addEventListener('DOMContentLoaded', () => {
        init();
        initExtractionModule();
    });
})();

