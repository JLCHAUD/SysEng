const { createApp, ref, reactive, computed, onMounted, watch, nextTick } = Vue;

// ── API helpers ───────────────────────────────────────────────────────────────
async function api(method, path, body) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (body !== undefined) opts.body = JSON.stringify(body);
  const r = await fetch(path, opts);
  if (!r.ok) {
    const err = await r.json().catch(() => ({ detail: r.statusText }));
    throw new Error(err.detail || r.statusText);
  }
  if (r.status === 204) return null;
  return r.json();
}
const GET  = path      => api('GET',    path);
const POST = (path, b) => api('POST',   path, b);
const PUT  = (path, b) => api('PUT',    path, b);
const DEL  = path      => api('DELETE', path);

// ── Toast (module-level) ──────────────────────────────────────────────────────
const toasts = ref([]);
let _toastId = 0;
function toast(msg, type = 'ok') {
  const id = ++_toastId;
  toasts.value.push({ id, msg, type });
  setTimeout(() => { toasts.value = toasts.value.filter(t => t.id !== id); }, 3500);
}
function toastErr(e) { toast(e instanceof Error ? e.message : String(e), 'error'); }

// ── Constants ─────────────────────────────────────────────────────────────────
// ROLES est maintenant chargé dynamiquement depuis /api/functions (functions.json partagé avec N1)
const PERIODICITES = ['quotidien','hebdomadaire','mensuel','manuel'];
const COL_TYPES    = ['string','float','int','date','pct','bool','KEY'];
const WRITE_MODES  = ['','creation','engineer','admin'];

const ROLE_BADGE = {
  pilote_pole:'badge-orange', pilote_metier:'badge-green',
  ingenieur_sys:'badge-blue', it_manager:'badge-purple',
  donneur_ordre_pole:'badge-orange', donneur_ordre_metier:'badge-gray',
  ingenieur_sys_client:'badge-green', expert_technique:'badge-purple',
  fournisseur:'badge-gray', correspondant_projet:'badge-gray',
};
const NODE_COLORS = {
  uo_instance:'#3b82f6', referentiel_uo:'#8b5cf6', referentiel_projet:'#ec4899',
  cockpit:'#f59e0b', pilote:'#10b981', consolidation:'#f97316', client:'#06b6d4',
  store:'#475569', unknown:'#64748b',
};

// ── Toast layer component ─────────────────────────────────────────────────────
const ToastLayer = {
  props: ['toasts'],
  template: `
    <div style="position:fixed;bottom:24px;right:24px;z-index:999;display:flex;flex-direction:column;gap:8px;">
      <transition-group name="toast">
        <div v-for="t in toasts" :key="t.id"
          :style="{background:t.type==='error'?'#7f1d1d':'#14532d',
                   border:'1px solid '+(t.type==='error'?'#ef4444':'#22c55e'),
                   padding:'10px 16px',borderRadius:'8px',fontSize:'0.83rem',
                   color:'#fff',maxWidth:'340px',boxShadow:'0 4px 12px rgba(0,0,0,0.5)'}">
          {{ t.msg }}
        </div>
      </transition-group>
    </div>`
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Dashboard
// ═══════════════════════════════════════════════════════════════════════════════
const ViewDashboard = {
  setup() {
    const status   = ref(null);
    const actors   = ref([]);
    const registry = ref([]);

    onMounted(async () => {
      try {
        [status.value, actors.value, registry.value] = await Promise.all([
          GET('/api/ecosystem/status').catch(() => null),
          GET('/api/actors'),
          GET('/api/registry'),
        ]);
      } catch(e) { toastErr(e); }
    });

    const syncOk  = computed(() => registry.value.filter(f => f.statut_dernier_synchro === 'ok').length);
    const syncErr = computed(() => registry.value.filter(f => f.statut_dernier_synchro && f.statut_dernier_synchro !== 'ok').length);
    const roleGroups = computed(() => {
      const m = {};
      actors.value.forEach(a => { m[a.role] = (m[a.role] || 0) + 1; });
      return Object.entries(m);
    });

    return { status, actors, registry, syncOk, syncErr, roleGroups, ROLE_BADGE };
  },
  template: `
    <div>
      <div class="stats-row">
        <div class="stat-card">
          <div class="stat-value" style="color:#10b981">{{ registry.length }}</div>
          <div class="stat-label">Posts enregistrés</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#60a5fa">{{ actors.length }}</div>
          <div class="stat-label">Acteurs</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#22c55e">{{ syncOk }}</div>
          <div class="stat-label">Syncs OK</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#f59e0b">{{ syncErr }}</div>
          <div class="stat-label">Syncs en erreur</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#94a3b8">{{ status ? status.edge_count : '…' }}</div>
          <div class="stat-label">Arêtes écosystème</div>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
        <div class="card">
          <div class="card-header"><span class="card-title">Posts récents</span></div>
          <table>
            <thead><tr><th>ID</th><th>Classe</th><th>Statut</th><th>Dernière sync</th></tr></thead>
            <tbody>
              <tr v-for="f in registry.slice(0,8)" :key="f.id">
                <td><strong>{{ f.id }}</strong></td>
                <td><span class="badge badge-teal" v-if="f.type_fichier">{{ f.type_fichier }}</span><span class="badge badge-gray" v-else>libre</span></td>
                <td>
                  <span class="badge" :class="f.statut_dernier_synchro==='ok'?'badge-green':'badge-orange'">
                    {{ f.statut_dernier_synchro || 'N/A' }}
                  </span>
                </td>
                <td style="color:var(--text-dim);font-size:0.76rem">{{ f.derniere_synchro ? f.derniere_synchro.slice(0,16) : '—' }}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="card">
          <div class="card-header"><span class="card-title">Acteurs par rôle</span></div>
          <table>
            <thead><tr><th>Rôle</th><th>Nb</th></tr></thead>
            <tbody>
              <tr v-for="[role, count] in roleGroups" :key="role">
                <td><span class="badge" :class="ROLE_BADGE[role]||'badge-gray'">{{ role }}</span></td>
                <td>{{ count }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Registre (Posts)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewRegistry = {
  setup() {
    const items     = ref([]);
    const fileTypes = ref([]);
    const showModal = ref(false);
    const editing   = ref(null);

    const blank = () => ({
      id: '', type_fichier: '', chemin: '', synchro_periodicite: 'quotidien',
      owner_role: '', genere_par_script: false
    });
    const form = reactive(blank());

    const load = async () => {
      try {
        [items.value, fileTypes.value] = await Promise.all([
          GET('/api/registry'), GET('/api/file-types')
        ]);
      } catch(e) { toastErr(e); }
    };
    onMounted(load);

    const openCreate = () => { Object.assign(form, blank()); editing.value = null; showModal.value = true; };
    const openEdit   = item => { Object.assign(form, { ...item }); editing.value = item.id; showModal.value = true; };

    const save = async () => {
      try {
        const payload = { ...form };
        if (!payload.type_fichier) payload.type_fichier = null;
        if (editing.value) {
          await PUT(`/api/registry/${editing.value}`, payload);
          toast('Post mis à jour');
        } else {
          await POST('/api/registry', payload);
          toast('Post ajouté');
        }
        showModal.value = false; await load();
      } catch(e) { toastErr(e); }
    };

    const del = async id => {
      if (!confirm(`Supprimer "${id}" du registre ?`)) return;
      try { await DEL(`/api/registry/${id}`); await load(); toast('Post supprimé'); }
      catch(e) { toastErr(e); }
    };

    return { items, fileTypes, showModal, editing, form, PERIODICITES,
             openCreate, openEdit, save, del };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <span class="card-title">Registre des Posts <span class="topbar-badge">{{ items.length }}</span></span>
          <button class="btn btn-primary" @click="openCreate">+ Ajouter un Post</button>
        </div>
        <table>
          <thead><tr><th>ID</th><th>Classe</th><th>Rôle owner</th><th>Périodicité</th><th>Statut</th><th>Chemin</th><th></th></tr></thead>
          <tbody>
            <tr v-for="f in items" :key="f.id">
              <td><strong>{{ f.id }}</strong></td>
              <td>
                <span class="badge badge-teal" v-if="f.type_fichier">{{ f.type_fichier }}</span>
                <span class="badge badge-gray" v-else>libre</span>
              </td>
              <td><span class="badge badge-gray" v-if="f.owner_role">{{ f.owner_role }}</span><span v-else style="color:var(--text-dim)">—</span></td>
              <td><span class="badge badge-gray">{{ f.synchro_periodicite }}</span></td>
              <td>
                <span class="badge" :class="f.statut_dernier_synchro==='ok'?'badge-green':'badge-orange'">
                  {{ f.statut_dernier_synchro || 'N/A' }}
                </span>
              </td>
              <td style="font-size:0.73rem;color:var(--text-dim);max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">{{ f.chemin }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(f)">Éditer</button>
                <button class="btn btn-danger btn-sm" @click="del(f.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Ajouter' }} un Post</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID *</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="UO-006" />
            </div>
            <div class="form-group">
              <label>Classe (optionnelle)</label>
              <select v-model="form.type_fichier">
                <option value="">— Post libre (sans Classe) —</option>
                <option v-for="t in fileTypes" :key="t.id" :value="t.id">{{ t.id }} — {{ t.label }}</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>Chemin *</label>
            <input v-model="form.chemin" placeholder="output/UOs/UO-006_xxx.xlsx" />
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Rôle owner</label>
              <input v-model="form.owner_role" placeholder="chef-projet" />
            </div>
            <div class="form-group">
              <label>Périodicité</label>
              <select v-model="form.synchro_periodicite">
                <option v-for="p in PERIODICITES" :key="p" :value="p">{{ p }}</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>
              <input type="checkbox" v-model="form.genere_par_script" style="width:auto;margin-right:6px" />
              Généré par script
            </label>
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">{{ editing ? 'Enregistrer' : 'Ajouter' }}</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Acteurs
// ═══════════════════════════════════════════════════════════════════════════════
const ViewActors = {
  setup() {
    const items     = ref([]);
    const functions = ref([]);   // chargé depuis /api/functions
    const showModal = ref(false);
    const editing   = ref(null);
    const blank = () => ({ id:'', nom:'', role:'', email:'' });
    const form    = reactive(blank());
    const formErr = ref('');

    const load = async () => {
      try {
        [items.value, functions.value] = await Promise.all([
          GET('/api/actors'),
          GET('/api/functions'),
        ]);
      } catch(e) { toastErr(e); }
    };
    onMounted(load);

    const roleLabel = id => {
      const f = functions.value.find(f => f.id === id);
      return f ? f.label : id;
    };
    const roleSide = id => {
      const f = functions.value.find(f => f.id === id);
      return f ? f.side : 'interne';
    };
    const roleBadgeClass = id => roleSide(id) === 'externe' ? 'badge-green' : (ROLE_BADGE[id] || 'badge-blue');

    const openCreate = async () => {
      Object.assign(form, blank());
      formErr.value = '';
      if (functions.value.length) form.role = functions.value[0].id;
      // suggère le prochain ID
      try { const r = await GET('/api/actors/next-id'); form.id = r.id; } catch(_) {}
      editing.value = null;
      showModal.value = true;
    };
    const openEdit = item => {
      Object.assign(form, { id: item.id, nom: item.nom, role: item.role, email: item.email || '' });
      formErr.value = '';
      editing.value = item.id;
      showModal.value = true;
    };

    const save = async () => {
      formErr.value = '';
      if (!form.id || !form.id.trim()) { formErr.value = "L'ID ne peut pas être vide."; return; }
      if (!form.nom || !form.nom.trim()) { formErr.value = "Le nom ne peut pas être vide."; return; }
      if (!editing.value && items.value.some(a => a.id === form.id.trim())) {
        formErr.value = `L'ID "${form.id.trim()}" est déjà utilisé.`; return;
      }
      try {
        if (editing.value) { await PUT(`/api/actors/${editing.value}`, form); toast('Acteur mis à jour'); }
        else               { await POST('/api/actors', form);                 toast('Acteur créé'); }
        showModal.value = false; await load();
      } catch(e) { formErr.value = e.message; toastErr(e); }
    };

    const del = async id => {
      if (!confirm(`Supprimer l'acteur "${id}" ?`)) return;
      try { await DEL(`/api/actors/${id}`); await load(); toast('Acteur supprimé'); }
      catch(e) { toastErr(e); }
    };

    const purgeEmpty = async () => {
      try { await DEL('/api/actors/purge-empty'); await load(); toast('Acteurs vides purgés'); }
      catch(e) { toastErr(e); }
    };

    const hasEmpty = computed(() => items.value.some(a => !a.id || !a.id.trim()));

    return { items, functions, showModal, editing, form, formErr, ROLE_BADGE,
             roleLabel, roleSide, roleBadgeClass, hasEmpty,
             openCreate, openEdit, save, del, purgeEmpty };
  },
  template: `
    <div>
      <!-- Bannière alerte IDs vides -->
      <div v-if="hasEmpty" style="margin-bottom:12px;padding:10px 16px;background:#7f1d1d;border-radius:8px;
                display:flex;align-items:center;justify-content:space-between;font-size:0.83rem;">
        <span style="color:#fca5a5;">⚠ Des acteurs avec un ID vide ont été détectés (données corrompues).</span>
        <button class="btn btn-danger btn-sm" @click="purgeEmpty">Purger les IDs vides</button>
      </div>

      <div class="card">
        <div class="card-header">
          <span class="card-title">Acteurs <span class="topbar-badge">{{ items.length }}</span></span>
          <button class="btn btn-primary" @click="openCreate">+ Nouvel acteur</button>
        </div>
        <table>
          <thead><tr><th>ID</th><th>Nom</th><th>Rôle</th><th>Email</th><th></th></tr></thead>
          <tbody>
            <tr v-for="a in items" :key="a.id">
              <td><code style="color:var(--accent);font-size:0.8rem">{{ a.id }}</code></td>
              <td><strong>{{ a.nom }}</strong></td>
              <td>
                <span class="badge" :class="roleBadgeClass(a.role)" :title="roleSide(a.role)">
                  {{ roleLabel(a.role) }}
                </span>
              </td>
              <td style="font-size:0.78rem;color:var(--text-dim)">{{ a.email }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(a)">Éditer</button>
                <button class="btn btn-danger btn-sm" @click="del(a.id)">✕</button>
              </td>
            </tr>
            <tr v-if="!items.length">
              <td colspan="5" style="text-align:center;color:var(--text-dim);padding:20px">
                Aucun acteur — cliquez sur "+ Nouvel acteur"
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvel' }} acteur</div>
          <div v-if="formErr" style="margin-bottom:12px;padding:8px 12px;background:#7f1d1d;
                border-radius:6px;font-size:0.82rem;color:#fca5a5;">⚠ {{ formErr }}</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID *</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="USR013" />
            </div>
            <div class="form-group">
              <label>Nom *</label>
              <input v-model="form.nom" placeholder="Prénom Nom" />
            </div>
          </div>
          <div class="form-group" style="margin-top:12px">
            <label>Rôle *</label>
            <select v-model="form.role">
              <optgroup label="Interne">
                <option v-for="f in functions.filter(f=>f.side==='interne')" :key="f.id" :value="f.id">
                  {{ f.label }}
                </option>
              </optgroup>
              <optgroup label="Externe">
                <option v-for="f in functions.filter(f=>f.side==='externe')" :key="f.id" :value="f.id">
                  {{ f.label }}
                </option>
              </optgroup>
            </select>
            <div v-if="!functions.length" style="font-size:0.75rem;color:#f87171;margin-top:4px;">
              ⚠ Aucun rôle défini — créez des rôles dans "Rôles & Fonctions"
            </div>
          </div>
          <div class="form-group">
            <label>Email</label>
            <input v-model="form.email" type="email" placeholder="prenom.nom@corp.fr" />
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">{{ editing ? 'Enregistrer' : 'Créer' }}</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Hiérarchie (LIST + COLLECT)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewHierarchy = {
  setup() {
    const lists     = ref([]);
    const collects  = ref([]);
    const registry  = ref([]);
    const activeTab = ref('lists');
    const showModal = ref(false);
    const modalType = ref('list');
    const editing   = ref(null);

    const blankList    = () => ({ id:'', owner_file_id:'', list_name:'', form:'TABLE', source_table:'', filter_type:'', filter_where:'' });
    const blankCollect = () => ({ id:'', owner_file_id:'', source_table:'', list_name:'', target_table:'', where_clause:'', cols_filter:'', with_fields:'' });
    const formL = reactive(blankList());
    const formC = reactive(blankCollect());

    const load = async () => {
      try {
        [lists.value, collects.value, registry.value] = await Promise.all([
          GET('/api/hierarchy/lists'), GET('/api/hierarchy/collects'), GET('/api/registry')
        ]);
      } catch(e) { toastErr(e); }
    };
    onMounted(load);

    const genId = () => Math.random().toString(36).slice(2, 10);

    const openCreateList    = () => { Object.assign(formL, blankList()); formL.id = genId(); editing.value = null; modalType.value = 'list';    showModal.value = true; };
    const openEditList      = item => { Object.assign(formL, { ...item }); editing.value = item.id; modalType.value = 'list';    showModal.value = true; };
    const openCreateCollect = () => { Object.assign(formC, blankCollect()); formC.id = genId(); editing.value = null; modalType.value = 'collect'; showModal.value = true; };
    const openEditCollect   = item => { Object.assign(formC, { ...item }); editing.value = item.id; modalType.value = 'collect'; showModal.value = true; };

    const saveList = async () => {
      try {
        if (editing.value) { await PUT(`/api/hierarchy/lists/${editing.value}`, formL); toast('LIST mise à jour'); }
        else { await POST('/api/hierarchy/lists', formL); toast('LIST créée'); }
        showModal.value = false; await load();
      } catch(e) { toastErr(e); }
    };
    const saveCollect = async () => {
      try {
        if (editing.value) { await PUT(`/api/hierarchy/collects/${editing.value}`, formC); toast('COLLECT mis à jour'); }
        else { await POST('/api/hierarchy/collects', formC); toast('COLLECT créé'); }
        showModal.value = false; await load();
      } catch(e) { toastErr(e); }
    };

    const delList    = async id => { if (!confirm('Supprimer cette LIST ?')) return; try { await DEL(`/api/hierarchy/lists/${id}`); await load(); toast('LIST supprimée'); } catch(e) { toastErr(e); } };
    const delCollect = async id => { if (!confirm('Supprimer ce COLLECT ?')) return; try { await DEL(`/api/hierarchy/collects/${id}`); await load(); toast('COLLECT supprimé'); } catch(e) { toastErr(e); } };

    return { lists, collects, registry, activeTab, showModal, modalType, editing,
             formL, formC, openCreateList, openEditList, openCreateCollect, openEditCollect,
             saveList, saveCollect, delList, delCollect };
  },
  template: `
    <div>
      <div style="display:flex;gap:0;margin-bottom:16px;border-bottom:1px solid var(--border)">
        <button class="btn btn-ghost" :style="activeTab==='lists'?'border-bottom:2px solid var(--accent);border-radius:0;color:var(--text)':'border-radius:0'" @click="activeTab='lists'">LIST <span class="topbar-badge">{{ lists.length }}</span></button>
        <button class="btn btn-ghost" :style="activeTab==='collects'?'border-bottom:2px solid var(--accent);border-radius:0;color:var(--text)':'border-radius:0'" @click="activeTab='collects'">COLLECT <span class="topbar-badge">{{ collects.length }}</span></button>
      </div>

      <div v-if="activeTab==='lists'">
        <div class="card">
          <div class="card-header">
            <span class="card-title">Déclarations LIST</span>
            <button class="btn btn-primary" @click="openCreateList">+ Nouvelle LIST</button>
          </div>
          <div v-if="lists.length===0" class="empty">Aucune LIST définie</div>
          <table v-else>
            <thead><tr><th>ID</th><th>Fichier père</th><th>Nom liste</th><th>Forme</th><th>Source table</th><th>Filtre</th><th></th></tr></thead>
            <tbody>
              <tr v-for="l in lists" :key="l.id">
                <td><code style="color:var(--accent);font-size:0.76rem">{{ l.id }}</code></td>
                <td><strong>{{ l.owner_file_id }}</strong></td>
                <td><span class="badge badge-purple">{{ l.list_name }}</span></td>
                <td><span class="badge" :class="l.form==='TABLE'?'badge-blue':'badge-orange'">{{ l.form }}</span></td>
                <td style="font-size:0.78rem;color:var(--text-dim)">{{ l.source_table || '—' }}</td>
                <td style="font-size:0.76rem;color:var(--text-dim)">{{ l.filter_type || '—' }}</td>
                <td style="display:flex;gap:6px;justify-content:flex-end">
                  <button class="btn btn-ghost btn-sm" @click="openEditList(l)">Éditer</button>
                  <button class="btn btn-danger btn-sm" @click="delList(l.id)">✕</button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <div v-if="activeTab==='collects'">
        <div class="card">
          <div class="card-header">
            <span class="card-title">Déclarations COLLECT</span>
            <button class="btn btn-primary" @click="openCreateCollect">+ Nouveau COLLECT</button>
          </div>
          <div v-if="collects.length===0" class="empty">Aucun COLLECT défini</div>
          <table v-else>
            <thead><tr><th>ID</th><th>Fichier père</th><th>Source table</th><th>From list</th><th>Into table</th><th>Where</th><th></th></tr></thead>
            <tbody>
              <tr v-for="c in collects" :key="c.id">
                <td><code style="color:var(--accent);font-size:0.76rem">{{ c.id }}</code></td>
                <td><strong>{{ c.owner_file_id }}</strong></td>
                <td><span class="badge badge-blue">{{ c.source_table }}</span></td>
                <td><span class="badge badge-purple">{{ c.list_name }}</span></td>
                <td><span class="badge badge-green">{{ c.target_table }}</span></td>
                <td style="font-size:0.76rem;color:var(--text-dim)">{{ c.where_clause || '—' }}</td>
                <td style="display:flex;gap:6px;justify-content:flex-end">
                  <button class="btn btn-ghost btn-sm" @click="openEditCollect(c)">Éditer</button>
                  <button class="btn btn-danger btn-sm" @click="delCollect(c.id)">✕</button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Modal LIST -->
      <div v-if="showModal && modalType==='list'" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvelle' }} LIST</div>
          <div class="form-row">
            <div class="form-group"><label>ID</label><input v-model="formL.id" :disabled="!!editing" /></div>
            <div class="form-group"><label>Forme</label><select v-model="formL.form"><option value="TABLE">TABLE</option><option value="DYNAMIC">DYNAMIC</option></select></div>
          </div>
          <div class="form-group">
            <label>Fichier père *</label>
            <select v-model="formL.owner_file_id">
              <option value="">— Choisir —</option>
              <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }}</option>
            </select>
          </div>
          <div class="form-group"><label>Nom de la liste *</label><input v-model="formL.list_name" placeholder="UOs_actifs" /></div>
          <div class="form-group"><label>Source table (DYNAMIC)</label><input v-model="formL.source_table" placeholder="TabUOs" /></div>
          <div class="form-row">
            <div class="form-group"><label>Type filtre</label><input v-model="formL.filter_type" placeholder="statut" /></div>
            <div class="form-group"><label>WHERE</label><input v-model="formL.filter_where" placeholder="actif" /></div>
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="saveList">{{ editing ? 'Enregistrer' : 'Créer' }}</button>
          </div>
        </div>
      </div>

      <!-- Modal COLLECT -->
      <div v-if="showModal && modalType==='collect'" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouveau' }} COLLECT</div>
          <div class="form-group">
            <label>Fichier père *</label>
            <select v-model="formC.owner_file_id">
              <option value="">— Choisir —</option>
              <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }}</option>
            </select>
          </div>
          <div class="form-row">
            <div class="form-group"><label>Table source *</label><input v-model="formC.source_table" placeholder="Planning" /></div>
            <div class="form-group"><label>From list *</label><input v-model="formC.list_name" placeholder="UOs_actifs" /></div>
          </div>
          <div class="form-group"><label>Into table *</label><input v-model="formC.target_table" placeholder="vue_planning" /></div>
          <div class="form-row">
            <div class="form-group"><label>WHERE</label><input v-model="formC.where_clause" placeholder="statut != fermé" /></div>
            <div class="form-group"><label>COLS filter</label><input v-model="formC.cols_filter" placeholder="nom,date,avancement" /></div>
          </div>
          <div class="form-group"><label>WITH fields</label><input v-model="formC.with_fields" placeholder="uo_id,uo_nom" /></div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="saveCollect">{{ editing ? 'Enregistrer' : 'Créer' }}</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Tables & colonnes
// ═══════════════════════════════════════════════════════════════════════════════
const ViewTables = {
  setup() {
    const tables       = ref([]);
    const registry     = ref([]);
    const selectedFile = ref('');
    const showModal    = ref(false);
    const editing      = ref(null);
    const importing    = ref(false);

    const blankTable = () => ({ id:'', file_id:'', table_name:'', sheet:'', description:'', columns:[] });
    const form = reactive(blankTable());

    const load = async () => {
      try { registry.value = await GET('/api/registry'); } catch(e) { toastErr(e); }
    };
    onMounted(load);

    const loadTables = async () => {
      if (!selectedFile.value) { tables.value = []; return; }
      tables.value = await GET(`/api/tables?file_id=${selectedFile.value}`);
    };
    watch(selectedFile, loadTables);

    const openCreate = () => {
      Object.assign(form, blankTable());
      form.file_id = selectedFile.value;
      form.id = (selectedFile.value || 'f') + '.' + Date.now().toString(36);
      editing.value = null; showModal.value = true;
    };
    const openEdit = item => { Object.assign(form, JSON.parse(JSON.stringify(item))); editing.value = item.id; showModal.value = true; };

    const addColumn    = () => form.columns.push({ name:'', col_type:'string', header:'', write:'', is_key:false, description:'' });
    const removeColumn = i  => form.columns.splice(i, 1);
    const moveUp   = i => { if (i > 0) { const t = form.columns[i]; form.columns[i] = form.columns[i-1]; form.columns[i-1] = t; } };
    const moveDown = i => { if (i < form.columns.length-1) { const t = form.columns[i]; form.columns[i] = form.columns[i+1]; form.columns[i+1] = t; } };

    const save = async () => {
      try {
        if (editing.value) { await PUT(`/api/tables/${editing.value}`, form); toast('Table mise à jour'); }
        else { await POST('/api/tables', form); toast('Table créée'); }
        showModal.value = false; await loadTables();
      } catch(e) { toastErr(e); }
    };
    const del = async id => {
      if (!confirm('Supprimer cette table ?')) return;
      try { await DEL(`/api/tables/${id}`); await loadTables(); toast('Table supprimée'); }
      catch(e) { toastErr(e); }
    };
    const importFromEco = async () => {
      if (!selectedFile.value) return;
      importing.value = true;
      try {
        const imported = await GET(`/api/tables/from-ecosystem/${selectedFile.value}`);
        let count = 0;
        for (const t of imported) { try { await POST('/api/tables', t); count++; } catch { /* already exists */ } }
        await loadTables();
        toast(`${count} table(s) importée(s) depuis l'écosystème`);
      } catch(e) { toastErr(e); }
      finally { importing.value = false; }
    };

    return { tables, registry, selectedFile, showModal, editing, importing, form,
             openCreate, openEdit, addColumn, removeColumn, moveUp, moveDown, save, del, importFromEco,
             COL_TYPES, WRITE_MODES };
  },
  template: `
    <div>
      <div class="card" style="margin-bottom:16px">
        <div style="display:flex;align-items:center;gap:12px">
          <label style="color:var(--text-dim);font-size:0.84rem;white-space:nowrap">Post :</label>
          <select v-model="selectedFile" style="flex:1;max-width:400px">
            <option value="">— Sélectionner un Post —</option>
            <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }}{{ f.type_fichier ? ' — ' + f.type_fichier : '' }}</option>
          </select>
          <button v-if="selectedFile" class="btn btn-ghost btn-sm" @click="importFromEco" :disabled="importing">
            {{ importing ? 'Import…' : '↓ Depuis écosystème' }}
          </button>
          <button v-if="selectedFile" class="btn btn-primary" @click="openCreate">+ Nouvelle table</button>
        </div>
      </div>

      <div v-if="!selectedFile" class="empty">Sélectionner un Post pour voir ses tables</div>

      <div v-for="tbl in tables" :key="tbl.id" class="card">
        <div class="card-header">
          <div>
            <span class="card-title">{{ tbl.table_name }}</span>
            <span style="color:var(--text-dim);font-size:0.78rem;margin-left:8px">SHEET={{ tbl.sheet }}</span>
            <span class="topbar-badge" style="margin-left:8px">{{ tbl.columns.length }} col.</span>
          </div>
          <div style="display:flex;gap:6px">
            <button class="btn btn-ghost btn-sm" @click="openEdit(tbl)">Éditer</button>
            <button class="btn btn-danger btn-sm" @click="del(tbl.id)">✕</button>
          </div>
        </div>
        <table>
          <thead><tr><th>Colonne</th><th>Type</th><th>Header Excel</th><th>Write mode</th><th>Clé</th></tr></thead>
          <tbody>
            <tr v-for="col in tbl.columns" :key="col.name">
              <td><code style="color:var(--accent)">{{ col.name }}</code></td>
              <td><span class="badge" :class="col.col_type==='KEY'?'badge-orange':['float','int'].includes(col.col_type)?'badge-blue':'badge-gray'">{{ col.col_type }}</span></td>
              <td style="color:var(--text-dim)">{{ col.header || '—' }}</td>
              <td style="font-size:0.76rem;color:var(--text-dim)">{{ col.write || '—' }}</td>
              <td>{{ col.is_key ? '✓' : '' }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal" style="width:700px;max-width:98vw">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvelle' }} table</div>
          <div class="form-row">
            <div class="form-group"><label>Nom de la table *</label><input v-model="form.table_name" placeholder="TabActivites" /></div>
            <div class="form-group"><label>Feuille Excel *</label><input v-model="form.sheet" placeholder="Activites" /></div>
          </div>
          <div class="form-group"><label>Description</label><input v-model="form.description" /></div>
          <div style="margin-top:16px;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between">
            <span style="font-size:0.84rem;font-weight:600">Colonnes</span>
            <button class="btn btn-ghost btn-sm" @click="addColumn">+ Colonne</button>
          </div>
          <div style="max-height:340px;overflow-y:auto">
            <table>
              <thead><tr><th style="width:28px"></th><th>Nom</th><th>Type</th><th>Header</th><th>Write</th><th>Clé</th><th style="width:36px"></th></tr></thead>
              <tbody>
                <tr v-for="(col,i) in form.columns" :key="i">
                  <td>
                    <div style="display:flex;flex-direction:column;gap:1px">
                      <span style="cursor:pointer;font-size:0.68rem;color:var(--text-dim)" @click="moveUp(i)">▲</span>
                      <span style="cursor:pointer;font-size:0.68rem;color:var(--text-dim)" @click="moveDown(i)">▼</span>
                    </div>
                  </td>
                  <td><input v-model="col.name" placeholder="nom" style="min-width:80px" /></td>
                  <td><select v-model="col.col_type" style="min-width:70px"><option v-for="t in COL_TYPES" :key="t" :value="t">{{ t }}</option></select></td>
                  <td><input v-model="col.header" placeholder="Libellé" /></td>
                  <td><select v-model="col.write" style="min-width:80px"><option v-for="w in WRITE_MODES" :key="w" :value="w">{{ w || '(none)' }}</option></select></td>
                  <td style="text-align:center"><input type="checkbox" v-model="col.is_key" style="width:auto" /></td>
                  <td><button class="btn btn-danger btn-sm" @click="removeColumn(i)">✕</button></td>
                </tr>
                <tr v-if="form.columns.length===0">
                  <td colspan="7" style="text-align:center;color:var(--text-dim);padding:16px">Aucune colonne</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">{{ editing ? 'Enregistrer' : 'Créer' }}</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Tissage (drag & drop PULL)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewTissage = {
  setup() {
    const registry      = ref([]);
    const tablesByPost  = ref({});   // { file_id: [TableDef] }
    const targetFileId  = ref('');
    const pulls         = ref([]);
    const saving        = ref(false);

    // Drag state
    const dragSrc = ref(null);   // { file_id, table_name, columns: [str] }
    const dropOver = ref(false);

    // Column selection step
    const pendingDrop    = ref(null);   // { file_id, table_name, columns }
    const selectedCols   = ref([]);
    const targetTable    = ref('');

    const mxlPreview = computed(() => {
      if (!pendingDrop.value || !targetFileId.value || !targetTable.value) return '';
      const cols = selectedCols.value.length ? `  COLS=${selectedCols.value.join(',')}` : '';
      return `PULL ${pendingDrop.value.file_id}::${pendingDrop.value.table_name} -> FILL_TABLE(?, ${targetTable.value})${cols}`;
    });

    const load = async () => {
      try {
        registry.value = await GET('/api/registry');
        pulls.value    = await GET('/api/hierarchy/pulls');
        // Load tables for all posts
        tablesByPost.value = {};
        await Promise.all(registry.value.map(async f => {
          const tbls = await GET(`/api/tables?file_id=${f.id}`).catch(() => []);
          if (tbls.length) tablesByPost.value[f.id] = tbls;
        }));
      } catch(e) { toastErr(e); }
    };
    onMounted(load);

    const onDragStart = (fileId, tbl, event) => {
      dragSrc.value = { file_id: fileId, table_name: tbl.table_name, columns: tbl.columns.map(c => c.name) };
      event.dataTransfer.effectAllowed = 'copy';
    };
    const onDragOver = event => { event.preventDefault(); dropOver.value = true; };
    const onDragLeave = () => { dropOver.value = false; };
    const onDrop = event => {
      event.preventDefault(); dropOver.value = false;
      if (!dragSrc.value || !targetFileId.value) return;
      pendingDrop.value  = { ...dragSrc.value };
      selectedCols.value = [...dragSrc.value.columns];
      targetTable.value  = dragSrc.value.table_name;
      dragSrc.value      = null;
    };

    const toggleCol = col => {
      const idx = selectedCols.value.indexOf(col);
      if (idx >= 0) selectedCols.value.splice(idx, 1);
      else selectedCols.value.push(col);
    };

    const declarePull = async () => {
      if (!pendingDrop.value || !targetFileId.value || !targetTable.value) return;
      saving.value = true;
      try {
        const p = await POST('/api/hierarchy/pulls', {
          source_file_id:  pendingDrop.value.file_id,
          source_table:    pendingDrop.value.table_name,
          target_file_id:  targetFileId.value,
          target_table:    targetTable.value,
          columns:         selectedCols.value,
        });
        pulls.value.push(p);
        toast(`Flux PULL déclaré : ${p.source_file_id}::${p.source_table} → ${p.target_file_id}::${p.target_table}`);
        pendingDrop.value = null; selectedCols.value = []; targetTable.value = '';
      } catch(e) { toastErr(e); }
      finally { saving.value = false; }
    };

    const cancelDrop = () => { pendingDrop.value = null; selectedCols.value = []; targetTable.value = ''; };

    const delPull = async id => {
      if (!confirm('Supprimer ce flux PULL ?')) return;
      try { await DEL(`/api/hierarchy/pulls/${id}`); pulls.value = pulls.value.filter(p => p.id !== id); toast('Flux PULL supprimé'); }
      catch(e) { toastErr(e); }
    };

    const sourcePosts = computed(() => registry.value.filter(f => tablesByPost.value[f.id]?.length && f.id !== targetFileId.value));

    return {
      registry, tablesByPost, targetFileId, pulls, saving,
      dropOver, pendingDrop, selectedCols, targetTable, mxlPreview,
      sourcePosts, onDragStart, onDragOver, onDragLeave, onDrop,
      toggleCol, declarePull, cancelDrop, delPull,
    };
  },
  template: `
    <div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">

        <!-- Panneau gauche : sources -->
        <div class="weave-panel">
          <div style="font-size:0.85rem;font-weight:600;margin-bottom:12px;color:var(--text-dim)">Posts sources — glisser une table vers la cible</div>
          <div v-if="Object.keys(tablesByPost).length===0" class="empty" style="padding:30px">Aucun Post avec tables définis</div>
          <div v-for="f in sourcePosts" :key="f.id" class="weave-post-block">
            <div class="weave-post-header">
              {{ f.id }}
              <span class="badge badge-teal" v-if="f.type_fichier" style="margin-left:6px">{{ f.type_fichier }}</span>
            </div>
            <div v-for="tbl in tablesByPost[f.id]" :key="tbl.id"
                 class="weave-table-item"
                 draggable="true"
                 @dragstart="onDragStart(f.id, tbl, $event)">
              ⇢ {{ tbl.table_name }}
              <span style="font-size:0.7rem;color:var(--text-dim);margin-left:4px">({{ tbl.columns.length }} cols)</span>
            </div>
          </div>
        </div>

        <!-- Panneau droit : cible -->
        <div class="weave-panel">
          <div style="font-size:0.85rem;font-weight:600;margin-bottom:12px;color:var(--text-dim)">Post cible</div>
          <div class="form-group" style="margin-bottom:14px">
            <select v-model="targetFileId" style="width:100%">
              <option value="">— Sélectionner le Post cible —</option>
              <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }}</option>
            </select>
          </div>

          <!-- Zone de drop -->
          <div v-if="!pendingDrop" class="drop-zone" :class="{over:dropOver}"
               @dragover="onDragOver" @dragleave="onDragLeave" @drop="onDrop">
            <div style="font-size:1.5rem;margin-bottom:8px">⇩</div>
            {{ targetFileId ? 'Déposer une table ici' : "Sélectionner un Post cible d'abord" }}
          </div>

          <!-- Panneau de configuration après drop -->
          <div v-if="pendingDrop" style="border:1px solid var(--accent);border-radius:8px;padding:16px">
            <div style="font-size:0.84rem;font-weight:600;margin-bottom:10px;color:var(--accent)">
              {{ pendingDrop.file_id }} :: {{ pendingDrop.table_name }}
            </div>

            <div class="form-group">
              <label>Nom de la table dans le Post cible</label>
              <input v-model="targetTable" :placeholder="pendingDrop.table_name" />
            </div>

            <div style="margin-bottom:12px">
              <div style="font-size:0.78rem;color:var(--text-dim);margin-bottom:6px">Colonnes à inclure :</div>
              <div style="display:flex;flex-wrap:wrap;gap:6px">
                <label v-for="col in pendingDrop.columns" :key="col"
                       style="display:flex;align-items:center;gap:4px;font-size:0.78rem;cursor:pointer;padding:3px 8px;border-radius:12px;border:1px solid var(--border)"
                       :style="{background:selectedCols.includes(col)?'rgba(16,185,129,0.15)':'transparent',borderColor:selectedCols.includes(col)?'var(--accent)':'var(--border)'}">
                  <input type="checkbox" :checked="selectedCols.includes(col)" @change="toggleCol(col)" style="width:auto;margin:0" />
                  {{ col }}
                </label>
              </div>
            </div>

            <div v-if="mxlPreview" style="background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:10px;margin-bottom:12px;font-family:monospace;font-size:0.8rem;color:#2dd4bf;word-break:break-all">
              {{ mxlPreview }}
            </div>

            <div style="display:flex;gap:8px">
              <button class="btn btn-ghost" @click="cancelDrop">Annuler</button>
              <button class="btn btn-primary" :disabled="!targetTable||saving" @click="declarePull">
                {{ saving ? '⏳…' : '⇢ Déclarer ce flux PULL' }}
              </button>
            </div>
          </div>
        </div>
      </div>

      <!-- Flux PULL déclarés -->
      <div class="card">
        <div class="card-header">
          <span class="card-title">Flux PULL déclarés <span class="topbar-badge">{{ pulls.length }}</span></span>
        </div>
        <div v-if="pulls.length===0" class="empty" style="padding:30px">Aucun flux PULL déclaré</div>
        <table v-else>
          <thead><tr><th>Source</th><th>Table source</th><th>→</th><th>Cible</th><th>Table cible</th><th>Colonnes</th><th></th></tr></thead>
          <tbody>
            <tr v-for="p in pulls" :key="p.id">
              <td><strong>{{ p.source_file_id }}</strong></td>
              <td><span class="badge badge-blue">{{ p.source_table }}</span></td>
              <td style="color:var(--accent);font-size:1rem">⇢</td>
              <td><strong>{{ p.target_file_id }}</strong></td>
              <td><span class="badge badge-green">{{ p.target_table }}</span></td>
              <td style="font-size:0.74rem;color:var(--text-dim)">{{ p.columns && p.columns.length ? p.columns.join(', ') : 'toutes' }}</td>
              <td><button class="btn btn-danger btn-sm" @click="delPull(p.id)">✕</button></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Générateur MXL
// ═══════════════════════════════════════════════════════════════════════════════
const ViewMxl = {
  setup() {
    const registry     = ref([]);
    const selectedFile = ref('');
    const mxl          = ref('');
    const loading      = ref(false);

    onMounted(async () => {
      try { registry.value = await GET('/api/registry'); } catch(e) { toastErr(e); }
    });

    const generate = async () => {
      if (!selectedFile.value) return;
      loading.value = true; mxl.value = '';
      try { const res = await GET(`/api/mxl/${selectedFile.value}`); mxl.value = res.mxl; }
      catch(e) { toastErr(e); }
      finally { loading.value = false; }
    };
    watch(selectedFile, generate);

    const copy = async () => { await navigator.clipboard.writeText(mxl.value); toast('MXL copié'); };
    const download = () => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([mxl.value], { type: 'text/plain' }));
      a.download = `manifeste_${selectedFile.value}.mxl`; a.click();
    };

    const highlighted = computed(() => {
      if (!mxl.value) return '';
      return mxl.value
        .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
        .replace(/^(FILE_TYPE|FILE_ID|VERSION)(\s+.*)$/gm, '<span style="color:#f97316;font-weight:600">$1</span><span style="color:#e2e8f0">$2</span>')
        .replace(/^(DEF|COL|PULL|PUSH|LIST|COLLECT|COMPUTE|BIND)(\s+)/gm, '<span style="color:#34d399;font-weight:600">$1</span><span style="color:#e2e8f0">$2</span>')
        .replace(/\b(TYPE|SHEET|HEADER|WRITE|KEY|FROM|TO|FORM|SOURCE|FILTER|WHERE|MODE|FROM_LIST|INTO|COLS|WITH)=/g, '<span style="color:#a78bfa">$1</span>=')
        .replace(/"([^"]*)"/g, '<span style="color:#60a5fa">"$1"</span>')
        .replace(/^(#.*)$/gm, '<span style="color:#475569;font-style:italic">$1</span>');
    });

    return { registry, selectedFile, mxl, loading, highlighted, generate, copy, download };
  },
  template: `
    <div>
      <div class="card" style="margin-bottom:16px">
        <div style="display:flex;align-items:center;gap:12px">
          <label style="color:var(--text-dim);font-size:0.84rem;white-space:nowrap">Post :</label>
          <select v-model="selectedFile" style="flex:1;max-width:400px">
            <option value="">— Sélectionner un Post —</option>
            <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }}</option>
          </select>
          <button class="btn btn-ghost btn-sm" @click="generate" :disabled="!selectedFile||loading">{{ loading ? '⌛…' : '↺ Regénérer' }}</button>
          <button v-if="mxl" class="btn btn-ghost btn-sm" @click="copy">⎘ Copier</button>
          <button v-if="mxl" class="btn btn-primary btn-sm" @click="download">↓ .mxl</button>
        </div>
      </div>
      <div v-if="!selectedFile" class="empty">Sélectionner un Post pour générer son Manifeste MXL</div>
      <div v-if="mxl" class="card" style="padding:0">
        <div style="padding:10px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between">
          <span style="font-size:0.78rem;color:var(--text-dim)">manifeste_{{ selectedFile }}.mxl</span>
          <span class="topbar-badge">{{ mxl.split('\\n').length }} lignes</span>
        </div>
        <pre style="padding:16px;font-size:0.82rem;line-height:1.6;overflow-x:auto;tab-size:2;font-family:'Cascadia Code','Fira Code',monospace"
             v-html="highlighted"></pre>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Import Excel → Post (wizard 5 étapes)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewExcelImport = {
  setup() {
    const step           = ref(1);
    const scanning       = ref(false);
    const saving         = ref(false);
    const comparing      = ref(false);
    const dragOver       = ref(false);
    const uploadedFile   = ref(null);
    const scanResult     = ref(null);
    const fileTypes      = ref([]);
    const availableTables = ref({});
    const selectedClass  = ref('');
    const tableComparisons = ref({});
    const tableMappings  = ref([]);
    const collectMappings = ref([]);
    const form = reactive({ file_id:'', owner_role:'', periodicite:'quotidien', update_class_fields:false });

    const isCockpit = computed(() => {
      const ft = fileTypes.value.find(f => f.id === selectedClass.value);
      return (ft?.allowed_namespaces || []).length > 0;
    });
    const childClasses = computed(() => {
      const ft = fileTypes.value.find(f => f.id === selectedClass.value);
      return ft?.allowed_namespaces || [];
    });
    const availableOwnerRoles = computed(() => {
      const roles = [...new Set(fileTypes.value.map(f => f.owner_role).filter(Boolean))];
      return roles.sort();
    });

    watch(selectedClass, cls => {
      const ft = fileTypes.value.find(f => f.id === cls);
      if (ft?.owner_role) form.owner_role = ft.owner_role;
    });

    onMounted(async () => {
      try { fileTypes.value = await GET('/api/file-types'); } catch(e) { toastErr(e); }
    });

    const onDrop = e => { dragOver.value = false; const f = e.dataTransfer?.files?.[0]; if (f) handleFile(f); };
    const onFileInput = e => { const f = e.target.files?.[0]; if (f) handleFile(f); };
    const handleFile = file => {
      if (!file.name.match(/\.(xlsx|xlsm|xlsb|xls)$/i)) { toast('Le fichier doit être un .xlsx / .xlsm', 'error'); return; }
      uploadedFile.value = file;
    };

    const scanFile = async () => {
      if (!uploadedFile.value) return;
      scanning.value = true;
      try {
        const fd = new FormData(); fd.append('file', uploadedFile.value);
        const r = await fetch('/api/excel/scan-upload', { method:'POST', body:fd });
        if (!r.ok) { const err = await r.json().catch(() => ({ detail: r.statusText })); throw new Error(err.detail || r.statusText); }
        scanResult.value = await r.json();
        const idRange = scanResult.value.named_ranges.find(nr => ['uo_id','file_id','post_id','id'].includes(nr.name.toLowerCase()));
        if (idRange?.value_preview) form.file_id = idRange.value_preview;
        step.value = 2;
      } catch(e) { toastErr(e); }
      finally { scanning.value = false; }
    };

    const compareAndProceed = async () => {
      if (!selectedClass.value && !confirm('Continuer sans Classe (Post libre) ?')) return;
      comparing.value = true;
      try {
        if (selectedClass.value && Object.keys(availableTables.value).length === 0) {
          const data = await GET('/api/excel/available-tables');
          availableTables.value = data.by_class || {};
        }
        const allExisting = [];
        Object.entries(availableTables.value).forEach(([cls, tbls]) => tbls.forEach(t => allExisting.push({ ...t, cls })));

        const comparisons = {};
        for (const tbl of (scanResult.value?.tables || [])) {
          let best = null, bestScore = 0;
          for (const existing of allExisting) {
            const res = await POST('/api/excel/compare-tables', { scanned_columns: tbl.columns.map(c => c.name), existing_table_id: existing.table_id });
            if (res.match_score > bestScore) { bestScore = res.match_score; best = { ...res, existing }; }
          }
          if (best && bestScore >= 0.6) comparisons[tbl.name] = best;
        }
        tableComparisons.value = comparisons;

        tableMappings.value = (scanResult.value?.tables || []).map(tbl => {
          const comp = comparisons[tbl.name];
          return {
            scannedTable: tbl,
            action: comp ? 'link' : 'create',
            linkedId: comp?.existing?.table_id || '',
            schemaName: comp?.existing?.table_name || tbl.name,
            columns: tbl.columns.map(c => ({ ...c, col_type: c.suggested_type, include: true, is_key: false, required: true })),
            comparison: comp || null,
          };
        });
        step.value = 3;
      } catch(e) { toastErr(e); }
      finally { comparing.value = false; }
    };

    const proceedFromMapping = async () => {
      if (isCockpit.value) {
        if (collectMappings.value.length === 0) {
          const suggestions = [];
          for (const ns of childClasses.value) {
            const nsKey = ns.replace(/\.$/, '');
            const cls = fileTypes.value.find(f => f.id.includes(nsKey) || (f.push_prefix || '').startsWith(nsKey));
            const classId = cls?.id || nsKey;
            const tbls = availableTables.value[classId] || [];
            tbls.forEach(t => suggestions.push({ source_class: classId, source_table: t.table_name, target_table: t.table_name, where:'', cols:'', enabled:false }));
          }
          collectMappings.value = suggestions;
        }
        step.value = 4;
      } else { step.value = 5; }
    };

    const register = async () => {
      if (!form.file_id) { toast('ID du Post requis', 'error'); return; }
      saving.value = true;
      try {
        const tblMappings = tableMappings.value.map(m => ({
          excel_name: m.scannedTable.name,
          schema_name: m.action === 'link' ? (m.linkedId?.split('.').pop() || m.schemaName) : m.schemaName,
          sheet: m.scannedTable.sheet,
          columns: m.columns.filter(c => c.include).map(c => ({ name:c.name, col_type:c.col_type, header:c.name, is_key:c.is_key, required:c.required })),
          include: true,
        }));
        const fieldMappings = (scanResult.value?.named_ranges || []).map(nr => ({
          named_range: nr.name, field_name: nr.name, field_type: nr.suggested_type,
          nature:'identitaire', source:'user_input', required:true, pushable:false, include:true,
        }));
        const activeCollects = collectMappings.value.filter(cm => cm.enabled).map(cm => ({
          source_class:cm.source_class, source_table:cm.source_table, target_table:cm.target_table,
          where_clause:cm.where, cols_filter:cm.cols,
        }));

        await POST('/api/excel/register', {
          path: uploadedFile.value?.name || '',
          file_id: form.file_id,
          type_fichier: selectedClass.value || null,
          owner_role: form.owner_role,
          synchro_periodicite: form.periodicite,
          field_mappings: fieldMappings, table_mappings: tblMappings,
          update_class_fields: form.update_class_fields, collect_mappings: activeCollects,
        });

        toast(`Post '${form.file_id}' enregistré avec succès`);
        step.value = 1; uploadedFile.value = null; scanResult.value = null;
        selectedClass.value = ''; tableMappings.value = []; collectMappings.value = [];
        tableComparisons.value = {}; availableTables.value = {};
        Object.assign(form, { file_id:'', owner_role:'', periodicite:'quotidien', update_class_fields:false });
      } catch(e) { toastErr(e); }
      finally { saving.value = false; }
    };

    const verdictClass = v => v === 'identical' ? 'badge-green' : v === 'compatible' ? 'badge-orange' : 'badge-gray';

    return {
      step, scanning, saving, comparing, dragOver, uploadedFile, scanResult,
      fileTypes, availableTables, selectedClass, tableComparisons,
      tableMappings, collectMappings, isCockpit, form, availableOwnerRoles,
      onDrop, onFileInput, handleFile, scanFile, compareAndProceed, proceedFromMapping,
      register, verdictClass,
    };
  },
  template: `
    <div>
      <!-- Indicateurs d'étapes -->
      <div style="display:flex;align-items:center;gap:4px;margin-bottom:20px;flex-wrap:wrap">
        <template v-for="(s, i) in ['Upload','Résultats','Mapping','COLLECT','Enregistrer']" :key="i">
          <div style="display:flex;align-items:center;gap:6px;padding:5px 12px;border-radius:6px;font-size:0.8rem;font-weight:500"
               :style="{background:step===i+1?'var(--accent)':'transparent',color:step===i+1?'#fff':'var(--text-dim)',opacity:(!isCockpit&&i===3)?0.3:1}">
            <span style="width:18px;height:18px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700"
                  :style="{background:step>i+1?'#22c55e':step===i+1?'rgba(255,255,255,0.25)':'var(--border)'}">
              {{ step > i+1 ? '✓' : i+1 }}
            </span>
            {{ s }}
          </div>
          <span v-if="i<4" style="color:var(--border)">›</span>
        </template>
      </div>

      <!-- ÉTAPE 1 : Upload -->
      <div v-if="step===1" class="card">
        <div class="card-header"><span class="card-title">Upload du fichier Excel</span></div>
        <div style="padding:16px">
          <div style="border:2px dashed var(--border);border-radius:10px;padding:40px;text-align:center;cursor:pointer;transition:border-color 0.15s"
               :style="{borderColor:dragOver?'var(--accent)':'var(--border)',background:dragOver?'rgba(16,185,129,0.04)':'transparent'}"
               @dragover.prevent="dragOver=true" @dragleave="dragOver=false" @drop.prevent="onDrop"
               @click="$refs.fileInput.click()">
            <input ref="fileInput" type="file" accept=".xlsx,.xlsm,.xlsb,.xls" style="display:none" @change="onFileInput" />
            <div style="font-size:2rem;margin-bottom:10px">📁</div>
            <div v-if="uploadedFile" style="color:var(--accent);font-weight:600">{{ uploadedFile.name }}</div>
            <div v-else style="color:var(--text-dim)">Glisser un fichier Excel ici, ou cliquer pour sélectionner</div>
          </div>
          <div style="margin-top:20px;text-align:right">
            <button class="btn btn-primary" :disabled="!uploadedFile||scanning" @click="scanFile">
              {{ scanning ? '⌛ Scan en cours…' : '→ Scanner le fichier' }}
            </button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 2 : Résultats du scan -->
      <div v-if="step===2" class="card">
        <div class="card-header"><span class="card-title">Résultats du scan</span></div>
        <div style="padding:16px">
          <div style="margin-bottom:16px">
            <div style="font-size:0.84rem;font-weight:600;margin-bottom:8px">Feuilles détectées :</div>
            <div style="display:flex;flex-wrap:wrap;gap:6px">
              <span v-for="s in scanResult.sheets" :key="s" class="badge badge-gray">{{ s }}</span>
            </div>
          </div>
          <div style="margin-bottom:16px">
            <div style="font-size:0.84rem;font-weight:600;margin-bottom:8px">Plages nommées ({{ scanResult.named_ranges.length }}) :</div>
            <div v-if="scanResult.named_ranges.length===0" style="color:var(--text-dim);font-size:0.82rem">Aucune plage nommée</div>
            <table v-else style="font-size:0.8rem">
              <thead><tr><th>Nom</th><th>Feuille</th><th>Type</th><th>Aperçu</th></tr></thead>
              <tbody>
                <tr v-for="nr in scanResult.named_ranges" :key="nr.name">
                  <td><code style="color:var(--accent)">{{ nr.name }}</code></td>
                  <td style="color:var(--text-dim)">{{ nr.sheet }}</td>
                  <td><span class="badge badge-gray">{{ nr.kind }}</span></td>
                  <td style="color:var(--text-dim)">{{ nr.value_preview || '—' }}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div style="margin-bottom:16px">
            <div style="font-size:0.84rem;font-weight:600;margin-bottom:8px">Tables Excel ({{ scanResult.tables.length }}) :</div>
            <div v-if="scanResult.tables.length===0" style="color:var(--text-dim);font-size:0.82rem">Aucune Table Excel trouvée</div>
            <table v-else style="font-size:0.8rem">
              <thead><tr><th>Nom</th><th>Feuille</th><th>Colonnes</th><th>Lignes</th></tr></thead>
              <tbody>
                <tr v-for="t in scanResult.tables" :key="t.name">
                  <td><strong>{{ t.name }}</strong></td>
                  <td style="color:var(--text-dim)">{{ t.sheet }}</td>
                  <td style="font-size:0.75rem;color:var(--text-dim)">{{ t.columns.map(c=>c.name).join(', ') }}</td>
                  <td>{{ t.row_count }}</td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="form-group">
            <label>Classe (optionnelle — laisser vide pour Post libre)</label>
            <select v-model="selectedClass">
              <option value="">— Post libre (sans Classe) —</option>
              <option v-for="ft in fileTypes" :key="ft.id" :value="ft.id">{{ ft.id }} — {{ ft.label }}</option>
            </select>
          </div>

          <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px">
            <button class="btn btn-ghost" @click="step=1">← Retour</button>
            <button class="btn btn-primary" :disabled="comparing" @click="compareAndProceed">
              {{ comparing ? '⌛ Analyse…' : '→ Mapping des tables' }}
            </button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 3 : Mapping des tables -->
      <div v-if="step===3" class="card">
        <div class="card-header"><span class="card-title">Mapping des tables</span></div>
        <div style="padding:16px">
          <div v-if="tableMappings.length===0" style="color:var(--text-dim);padding:20px;text-align:center">Aucune Table Excel dans ce fichier</div>
          <div v-for="m in tableMappings" :key="m.scannedTable.name" style="border:1px solid var(--border);border-radius:8px;padding:14px;margin-bottom:12px">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
              <strong style="font-size:0.88rem">{{ m.scannedTable.name }}</strong>
              <span style="color:var(--text-dim);font-size:0.78rem">{{ m.scannedTable.sheet }}</span>
              <div style="margin-left:auto;display:flex;align-items:center;gap:8px">
                <select v-model="m.action" style="width:120px">
                  <option value="create">Créer</option>
                  <option value="link">Lier à existante</option>
                  <option value="skip">Ignorer</option>
                </select>
                <input v-if="m.action!=='skip'" v-model="m.schemaName" style="width:180px" placeholder="Nom dans schéma" />
              </div>
            </div>
            <div v-if="m.comparison" style="font-size:0.78rem;color:var(--text-dim);margin-bottom:8px">
              Similarité : <span :class="'badge '+verdictClass(m.comparison.verdict)">{{ m.comparison.verdict }} ({{ Math.round(m.comparison.match_score*100) }}%)</span>
            </div>
            <div v-if="m.action!=='skip'" style="display:flex;flex-wrap:wrap;gap:4px">
              <label v-for="col in m.columns" :key="col.name"
                     style="display:flex;align-items:center;gap:4px;font-size:0.76rem;padding:2px 7px;border-radius:10px;cursor:pointer;border:1px solid var(--border)"
                     :style="{opacity:col.include?1:0.4}">
                <input type="checkbox" v-model="col.include" style="width:auto;margin:0" />
                {{ col.name }} <span style="color:var(--text-dim);font-size:0.7rem">({{ col.col_type }})</span>
              </label>
            </div>
          </div>
          <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px">
            <button class="btn btn-ghost" @click="step=2">← Retour</button>
            <button class="btn btn-primary" @click="proceedFromMapping">→ Suivant</button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 4 : COLLECT (cockpits uniquement) -->
      <div v-if="step===4" class="card">
        <div class="card-header"><span class="card-title">Déclarations COLLECT</span></div>
        <div style="padding:16px">
          <div v-if="collectMappings.length===0" style="color:var(--text-dim);padding:20px;text-align:center">Aucune suggestion COLLECT disponible</div>
          <table v-else style="font-size:0.82rem">
            <thead><tr><th style="width:28px"></th><th>Source (classe)</th><th>Table source</th><th>Dans ce Post comme</th><th>WHERE</th></tr></thead>
            <tbody>
              <tr v-for="cm in collectMappings" :key="cm.source_class+'_'+cm.source_table">
                <td><input type="checkbox" v-model="cm.enabled" style="width:auto" /></td>
                <td><span class="badge badge-gray" style="font-size:0.73rem">{{ cm.source_class }}</span></td>
                <td><code style="color:var(--accent)">{{ cm.source_table }}</code></td>
                <td><input v-model="cm.target_table" :disabled="!cm.enabled" style="width:160px" /></td>
                <td><input v-model="cm.where" :disabled="!cm.enabled" placeholder="statut = 'actif'" style="width:180px" /></td>
              </tr>
            </tbody>
          </table>
          <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px">
            <button class="btn btn-ghost" @click="step=3">← Retour</button>
            <button class="btn btn-primary" @click="step=5">→ Enregistrement</button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 5 : Enregistrement -->
      <div v-if="step===5" class="card">
        <div class="card-header"><span class="card-title">Enregistrement du Post</span></div>
        <div style="padding:16px">
          <div class="form-row">
            <div class="form-group"><label>ID du Post *</label><input v-model="form.file_id" placeholder="UO-006" /></div>
            <div class="form-group"><label>Classe</label><input :value="selectedClass||'(Post libre)'" disabled /></div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Rôle owner</label>
              <select v-model="form.owner_role">
                <option value="">— Choisir —</option>
                <option v-for="r in availableOwnerRoles" :key="r" :value="r">{{ r }}</option>
              </select>
            </div>
            <div class="form-group">
              <label>Périodicité</label>
              <select v-model="form.periodicite">
                <option value="quotidien">quotidien</option>
                <option value="hebdomadaire">hebdomadaire</option>
                <option value="mensuel">mensuel</option>
                <option value="manuel">manuel</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label><input type="checkbox" v-model="form.update_class_fields" style="width:auto;margin-right:6px" />Mettre à jour les min_fields de la Classe</label>
          </div>

          <div style="background:var(--surface-alt);border:1px solid var(--border);border-radius:8px;padding:14px;margin:16px 0;font-size:0.84rem;line-height:1.9">
            <div style="font-weight:600;margin-bottom:6px">Résumé</div>
            <div style="color:#22c55e">✓ {{ (scanResult?.named_ranges||[]).length }} plages nommées</div>
            <div style="color:#22c55e">✓ {{ tableMappings.filter(m=>m.action==='link').length }} table(s) liée(s)</div>
            <div style="color:#60a5fa">✓ {{ tableMappings.filter(m=>m.action==='create').length }} nouvelle(s) table(s)</div>
            <div v-if="isCockpit" style="color:#c084fc">✓ {{ collectMappings.filter(cm=>cm.enabled).length }} COLLECT</div>
          </div>

          <div style="display:flex;gap:8px;justify-content:space-between">
            <button class="btn btn-ghost" @click="step=isCockpit?4:3">← Retour</button>
            <button class="btn btn-primary" :disabled="!form.file_id||saving" @click="register">
              {{ saving ? '⏳ Enregistrement…' : '✓ Enregistrer le Post' }}
            </button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Graphe écosystème (Cytoscape.js)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewEcosystem = {
  setup() {
    const graphData   = ref(null);
    const status      = ref(null);
    const cyInstance  = ref(null);
    const detail      = ref(null);

    onMounted(async () => {
      try {
        [graphData.value, status.value] = await Promise.all([
          GET('/api/ecosystem/graph'), GET('/api/ecosystem/status')
        ]);
        await nextTick();
        buildGraph();
      } catch(e) { toastErr(e); }
    });

    const resolveNodeId = node => {
      if (!node) return null;
      if (node.startsWith('store::')) return 'store_' + node.replace('store::', '').split('.')[0];
      return node.split('::')[0];
    };

    const buildGraph = () => {
      if (!graphData.value || typeof cytoscape === 'undefined') return;
      const elements = [];

      graphData.value.nodes.forEach(n => {
        elements.push({ data: { id:n.id, label:n.id, type:n.file_type, status:n.status, path:n.path, color:NODE_COLORS[n.file_type]||NODE_COLORS.unknown } });
      });

      const storeGroups = {};
      graphData.value.edges.forEach(e => {
        [e.from_node, e.to_node].forEach(n => {
          if (n.startsWith('store::')) storeGroups[n.replace('store::', '').split('.')[0]] = true;
        });
      });
      Object.keys(storeGroups).forEach(k => {
        elements.push({ data: { id:'store_'+k, label:'store::\n'+k+'.*', type:'store', color:NODE_COLORS.store } });
      });

      const edgeMap = {};
      graphData.value.edges.forEach(e => {
        const from = resolveNodeId(e.from_node);
        const to   = resolveNodeId(e.to_node);
        if (!from || !to || from === to) return;
        const key = `${from}→${to}→${e.edge_type}`;
        edgeMap[key] = { from, to, type: e.edge_type };
      });
      Object.entries(edgeMap).forEach(([k, e]) => {
        elements.push({ data: { id:'e_'+k, source:e.from, target:e.to, edgeType:e.type,
          color: e.type==='PUSH'?'#22c55e':e.type==='PULL'?'#60a5fa':'#f97316' } });
      });

      const cy = cytoscape({
        container: document.getElementById('cy'), elements,
        style: [
          { selector:'node', style: { 'background-color':'data(color)', 'label':'data(label)', 'color':'#e2e8f0', 'font-size':'11px', 'text-valign':'bottom', 'text-halign':'center', 'text-margin-y':'4px', width:46, height:46, 'border-width':2, 'border-color':'#1e293b', 'text-wrap':'wrap', 'text-max-width':'80px' } },
          { selector:'node[type="store"]', style:{ shape:'rectangle', width:60, height:36, 'font-size':'10px' } },
          { selector:'edge', style:{ 'line-color':'data(color)', 'target-arrow-color':'data(color)', 'target-arrow-shape':'triangle', 'curve-style':'bezier', width:2, label:'data(edgeType)', 'font-size':'9px', color:'#94a3b8', 'text-rotation':'autorotate' } },
          { selector:':selected', style:{ 'border-width':3, 'border-color':'#fff' } },
          { selector:'.faded', style:{ opacity:0.15 } },
        ],
        layout:{ name:'cose', idealEdgeLength:140, nodeOverlap:20, refresh:20, fit:true, padding:40, randomize:false, animate:false }
      });

      cy.on('tap', 'node', evt => {
        detail.value = evt.target.data();
        cy.elements().addClass('faded');
        evt.target.removeClass('faded');
        evt.target.neighborhood().removeClass('faded');
      });
      cy.on('tap', evt => { if (evt.target === cy) { detail.value = null; cy.elements().removeClass('faded'); } });
      cyInstance.value = cy;
    };

    const relayout = () => { if (cyInstance.value) cyInstance.value.layout({ name:'cose', animate:false, fit:true, padding:40 }).run(); };
    const fitGraph  = () => { if (cyInstance.value) cyInstance.value.fit(40); };

    return { graphData, status, detail, relayout, fitGraph };
  },
  template: `
    <div>
      <div v-if="status" style="display:flex;gap:12px;margin-bottom:16px;align-items:center">
        <div class="stat-card" style="padding:10px 18px;display:flex;align-items:center;gap:10px">
          <div style="font-size:1.3rem;font-weight:700;color:#22c55e">{{ status.ok }}</div>
          <div style="font-size:0.73rem;color:var(--text-dim)">fichiers OK</div>
        </div>
        <div class="stat-card" style="padding:10px 18px;display:flex;align-items:center;gap:10px">
          <div style="font-size:1.3rem;font-weight:700;color:#f59e0b">{{ status.edge_count }}</div>
          <div style="font-size:0.73rem;color:var(--text-dim)">arêtes</div>
        </div>
        <div style="font-size:0.76rem;color:var(--text-dim)">Dernier scan : {{ status.last_scan ? status.last_scan.slice(0,16) : '—' }}</div>
        <div style="flex:1"></div>
        <button class="btn btn-ghost btn-sm" @click="fitGraph">⊡ Recadrer</button>
        <button class="btn btn-ghost btn-sm" @click="relayout">↺ Relayout</button>
      </div>

      <div style="position:relative">
        <div id="cy"></div>
        <div v-if="detail" style="position:absolute;top:12px;right:12px;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px;min-width:200px;font-size:0.8rem">
          <div style="font-weight:600;margin-bottom:8px">{{ detail.label }}</div>
          <div v-if="detail.type !== 'store'">
            <div style="color:var(--text-dim)">Type</div>
            <div class="badge badge-teal" style="margin-bottom:8px">{{ detail.type }}</div><br>
            <div style="color:var(--text-dim)">Statut</div>
            <div class="badge" :class="detail.status==='ok'?'badge-green':'badge-orange'">{{ detail.status }}</div><br><br>
            <div style="color:var(--text-dim);font-size:0.73rem;word-break:break-all">{{ detail.path }}</div>
          </div>
          <div v-else style="color:var(--text-dim)">Nœud store central</div>
        </div>
      </div>

      <div style="display:flex;gap:12px;margin-top:10px;font-size:0.76rem;color:var(--text-dim);flex-wrap:wrap">
        <span style="display:flex;align-items:center;gap:4px"><span style="width:20px;height:3px;background:#22c55e;display:inline-block"></span> PUSH</span>
        <span style="display:flex;align-items:center;gap:4px"><span style="width:20px;height:3px;background:#60a5fa;display:inline-block"></span> PULL</span>
        <span style="display:flex;align-items:center;gap:4px"><span style="width:20px;height:3px;background:#f97316;display:inline-block"></span> COLLECT</span>
        <span style="margin-left:12px">Cliquer un nœud pour voir ses connexions</span>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Ecosystem Manager (Espace de travail)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewEcosystemManager = {
  setup() {
    const ecosystems = ref([]);
    const current = ref(null);
    const showNew = ref(false);
    const newForm = ref({ name: '', path: '' });
    const loading = ref(false);

    async function load() {
      try {
        [ecosystems.value, current.value] = await Promise.all([
          GET('/api/ecosystem/list'),
          GET('/api/ecosystem/current'),
        ]);
      } catch(e) { toastErr(e); }
    }

    async function activate(path) {
      try {
        await POST('/api/ecosystem/open', { path });
        await load();
        toast('Écosystème activé');
      } catch(e) { toastErr(e); }
    }

    async function createNew() {
      if (!newForm.value.name || !newForm.value.path) return;
      loading.value = true;
      try {
        await POST('/api/ecosystem/new', newForm.value);
        newForm.value = { name: '', path: '' };
        showNew.value = false;
        await load();
        toast('Écosystème créé et activé');
      } catch(e) { toastErr(e); }
      finally { loading.value = false; }
    }

    async function remove(name) {
      if (!confirm(`Retirer "${name}" de la liste ?`)) return;
      try {
        await DEL(`/api/ecosystem/${encodeURIComponent(name)}`);
        await load();
        toast('Écosystème retiré');
      } catch(e) { toastErr(e); }
    }

    onMounted(load);
    return { ecosystems, current, showNew, newForm, loading, activate, createNew, remove };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <span class="card-title">Espace de travail actif</span>
        </div>
        <div v-if="current" style="display:flex;align-items:center;gap:12px;padding:8px 0;">
          <span style="font-size:1.5rem;">◎</span>
          <div>
            <div style="font-weight:600;">{{ current.name }}</div>
            <div style="font-size:0.75rem;color:var(--text-dim);">{{ current.path }}</div>
          </div>
          <span class="badge" :class="current.valid ? 'badge-green' : 'badge-orange'" style="margin-left:auto;">
            {{ current.valid ? 'valide' : 'invalide' }}
          </span>
          <span class="badge badge-blue">{{ current.post_count }} posts</span>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <span class="card-title">Écosystèmes connus</span>
          <button class="btn btn-primary btn-sm" @click="showNew=!showNew">+ Nouveau</button>
        </div>

        <div v-if="showNew" style="background:var(--surface2);border-radius:8px;padding:16px;margin-bottom:16px;">
          <div class="form-row" style="margin-bottom:12px;">
            <div class="form-group" style="margin:0;">
              <label>Nom</label>
              <input v-model="newForm.name" placeholder="Mon écosystème" />
            </div>
            <div class="form-group" style="margin:0;">
              <label>Chemin (absolu ou relatif au projet)</label>
              <input v-model="newForm.path" placeholder="config_projet2" />
            </div>
          </div>
          <div style="display:flex;gap:8px;justify-content:flex-end;">
            <button class="btn btn-ghost btn-sm" @click="showNew=false">Annuler</button>
            <button class="btn btn-primary btn-sm" :disabled="loading" @click="createNew">Créer et activer</button>
          </div>
        </div>

        <table>
          <thead><tr>
            <th>Nom</th><th>Chemin</th><th>Posts</th><th>État</th><th></th>
          </tr></thead>
          <tbody>
            <tr v-for="eco in ecosystems" :key="eco.path">
              <td>
                <span style="font-weight:500;">{{ eco.name }}</span>
                <span v-if="eco.active" class="badge badge-blue" style="margin-left:8px;">actif</span>
              </td>
              <td style="font-size:0.75rem;color:var(--text-dim);">{{ eco.path }}</td>
              <td>{{ eco.post_count }}</td>
              <td><span class="badge" :class="eco.valid ? 'badge-green' : 'badge-orange'">{{ eco.valid ? 'valide' : 'invalide' }}</span></td>
              <td style="text-align:right;display:flex;gap:6px;justify-content:flex-end;">
                <button v-if="!eco.active" class="btn btn-ghost btn-sm" @click="activate(eco.path)">Activer</button>
                <button v-if="!eco.active" class="btn btn-danger btn-sm" @click="remove(eco.name)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>`
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Synchronisation
// ═══════════════════════════════════════════════════════════════════════════════
const ViewSync = {
  setup() {
    const posts      = ref([]);
    const selected   = ref([]);        // IDs cochés
    const running    = ref(false);
    const rapport    = ref(null);
    const reports    = ref([]);
    const expanded   = ref({});        // logs dépliés par ID
    const activeReport = ref(null);    // rapport historique ouvert

    async function load() {
      try {
        const [reg, status] = await Promise.all([
          GET('/api/registry'),
          GET('/api/sync/status'),
        ]);
        posts.value = reg;
        running.value = status.running;
        if (status.last_rapport) rapport.value = status.last_rapport;
        reports.value = await GET('/api/sync/reports');
      } catch(e) { toastErr(e); }
    }

    function toggleAll() {
      if (selected.value.length === posts.value.length) {
        selected.value = [];
      } else {
        selected.value = posts.value.map(p => p.id);
      }
    }

    function fileStatut(id) {
      if (!rapport.value) return null;
      return rapport.value.fichiers?.find(f => f.id === id);
    }

    async function lancer(idsOverride) {
      const ids = idsOverride ?? (selected.value.length ? selected.value : []);
      running.value = true;
      rapport.value = null;
      try {
        rapport.value = await POST('/api/sync/run', { ids, force: false });
        toast(`Sync terminé — ${rapport.value.nb_ok} OK / ${rapport.value.nb_erreur} erreurs`);
        await load();
      } catch(e) {
        toastErr(e);
        running.value = false;
      }
    }

    async function openHistorique(filename) {
      try { activeReport.value = await GET(`/api/sync/reports/${filename}`); }
      catch(e) { toastErr(e); }
    }

    function statutClass(s) {
      if (s === 'ok') return 'badge-green';
      if (s === 'skip_verrouille') return 'badge-orange';
      if (s === 'erreur') return 'badge-red';
      return 'badge-gray';
    }

    onMounted(load);
    return {
      posts, selected, running, rapport, reports, expanded, activeReport,
      toggleAll, fileStatut, lancer, openHistorique, statutClass,
    };
  },
  template: `
    <div>
      <!-- Barre d'actions -->
      <div class="card">
        <div class="card-header">
          <span class="card-title">Synchronisation ExoSync</span>
          <div style="display:flex;gap:8px;align-items:center;">
            <span v-if="running" style="color:var(--accent);font-size:0.85rem;">⟳ Sync en cours…</span>
            <button class="btn btn-ghost btn-sm" @click="lancer([])" :disabled="running">
              ↺ Tout synchroniser
            </button>
            <button class="btn btn-primary btn-sm"
                    @click="lancer(selected)" :disabled="running || !selected.length">
              ▶ Sync sélection ({{ selected.length }})
            </button>
          </div>
        </div>

        <!-- Tableau des Posts -->
        <table>
          <thead><tr>
            <th style="width:32px;"><input type="checkbox" @change="toggleAll"
                :checked="selected.length === posts.length && posts.length > 0" /></th>
            <th>ID Post</th><th>Classe</th><th>Dernière sync</th><th>Statut</th>
          </tr></thead>
          <tbody>
            <tr v-if="!posts.length">
              <td colspan="5" style="text-align:center;color:var(--text-dim);padding:24px;">
                Registre vide — ajoutez des Posts avant de synchroniser.
              </td>
            </tr>
            <tr v-for="p in posts" :key="p.id">
              <td><input type="checkbox" :value="p.id" v-model="selected" /></td>
              <td><strong>{{ p.id }}</strong></td>
              <td>
                <span class="badge badge-teal" v-if="p.type_fichier">{{ p.type_fichier }}</span>
                <span class="badge badge-gray" v-else>libre</span>
              </td>
              <td style="font-size:0.78rem;color:var(--text-dim);">
                {{ p.derniere_synchro ? p.derniere_synchro.replace('T',' ') : '—' }}
              </td>
              <td>
                <span v-if="fileStatut(p.id)" class="badge"
                      :class="statutClass(fileStatut(p.id).statut)">
                  {{ fileStatut(p.id).statut }}
                </span>
                <span v-else-if="p.statut_dernier_synchro" class="badge"
                      :class="statutClass(p.statut_dernier_synchro)">
                  {{ p.statut_dernier_synchro }}
                </span>
                <span v-else style="color:var(--text-dim);">—</span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Résultat du dernier sync -->
      <div v-if="rapport" class="card">
        <div class="card-header">
          <span class="card-title">Résultat — {{ rapport.debut?.replace('T',' ') }}</span>
          <div style="display:flex;gap:8px;">
            <span class="badge badge-green">{{ rapport.nb_ok }} OK</span>
            <span v-if="rapport.nb_skip" class="badge badge-orange">{{ rapport.nb_skip }} skip</span>
            <span v-if="rapport.nb_erreur" class="badge badge-red">{{ rapport.nb_erreur }} erreur(s)</span>
          </div>
        </div>
        <div v-for="f in rapport.fichiers" :key="f.id"
             style="border-bottom:1px solid var(--border);padding:8px 0;">
          <div style="display:flex;align-items:center;gap:8px;cursor:pointer;"
               @click="expanded[f.id] = !expanded[f.id]">
            <span class="badge" :class="statutClass(f.statut)" style="min-width:80px;text-align:center;">
              {{ f.statut }}
            </span>
            <strong>{{ f.id }}</strong>
            <span style="font-size:0.78rem;color:var(--text-dim);">{{ f.chemin }}</span>
            <span style="margin-left:auto;font-size:0.75rem;color:var(--text-dim);">
              {{ expanded[f.id] ? '▲' : '▼' }} logs
            </span>
          </div>
          <div v-if="expanded[f.id]"
               style="margin-top:8px;background:var(--surface2);border-radius:6px;padding:8px;font-size:0.75rem;font-family:monospace;white-space:pre-wrap;line-height:1.6;">
            {{ f.log.join('\\n') || '(aucun log)' }}
          </div>
        </div>
      </div>

      <!-- Historique des rapports -->
      <div class="card">
        <div class="card-header">
          <span class="card-title">Historique des synchronisations</span>
        </div>
        <table>
          <thead><tr><th>Date</th><th>Total</th><th>OK</th><th>Erreurs</th><th></th></tr></thead>
          <tbody>
            <tr v-if="!reports.length">
              <td colspan="5" style="text-align:center;color:var(--text-dim);padding:16px;">
                Aucun rapport disponible.
              </td>
            </tr>
            <tr v-for="r in reports" :key="r.filename">
              <td style="font-size:0.82rem;">{{ r.debut ? r.debut.replace('T',' ') : r.filename }}</td>
              <td>{{ r.nb_total }}</td>
              <td><span class="badge badge-green">{{ r.nb_ok }}</span></td>
              <td>
                <span class="badge" :class="r.nb_erreur ? 'badge-red' : 'badge-gray'">
                  {{ r.nb_erreur }}
                </span>
              </td>
              <td style="text-align:right;">
                <button class="btn btn-ghost btn-sm" @click="openHistorique(r.filename)">Voir</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Modal rapport historique -->
      <div v-if="activeReport" class="modal-overlay" @click.self="activeReport=null">
        <div class="modal" style="max-width:700px;max-height:80vh;overflow-y:auto;">
          <div class="modal-title">Rapport — {{ activeReport.debut?.replace('T',' ') }}</div>
          <div style="display:flex;gap:8px;margin-bottom:16px;">
            <span class="badge badge-green">{{ activeReport.nb_ok }} OK</span>
            <span v-if="activeReport.nb_skip" class="badge badge-orange">{{ activeReport.nb_skip }} skip</span>
            <span v-if="activeReport.nb_erreur" class="badge badge-red">{{ activeReport.nb_erreur }} erreur(s)</span>
          </div>
          <div v-for="f in activeReport.fichiers" :key="f.id"
               style="margin-bottom:12px;border:1px solid var(--border);border-radius:6px;padding:10px;">
            <div style="display:flex;gap:8px;align-items:center;margin-bottom:6px;">
              <span class="badge" :class="statutClass(f.statut)">{{ f.statut }}</span>
              <strong>{{ f.id }}</strong>
              <span style="font-size:0.75rem;color:var(--text-dim);">{{ f.chemin }}</span>
            </div>
            <div style="background:var(--surface2);border-radius:4px;padding:8px;font-size:0.73rem;font-family:monospace;white-space:pre-wrap;">
              {{ f.log.join('\\n') || '(aucun log)' }}
            </div>
          </div>
          <div style="text-align:right;margin-top:12px;">
            <button class="btn btn-ghost" @click="activeReport=null">Fermer</button>
          </div>
        </div>
      </div>
    </div>`
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Gabarits & Affaires
// ═══════════════════════════════════════════════════════════════════════════════
const ViewGabarits = {
  setup() {
    const gabarits    = ref([]);
    const showNew     = ref(false);
    const showOpen    = ref(false);
    const showClone   = ref(null);   // gabarit source sélectionné
    const showExport  = ref(false);
    const loading     = ref(false);

    const newForm    = reactive({ name: '', path: '', description: '' });
    const openForm   = reactive({ path: '', name: '', description: '' });
    const cloneForm  = reactive({ dest_path: '', name: '', activate: true });
    const exportForm = reactive({ name: '', dest_path: '', description: '' });

    async function load() {
      try { gabarits.value = await GET('/api/gabarits/list'); }
      catch(e) { toastErr(e); }
    }

    async function createNew() {
      if (!newForm.name || !newForm.path) return;
      loading.value = true;
      try {
        await POST('/api/gabarits/new', { ...newForm });
        Object.assign(newForm, { name: '', path: '', description: '' });
        showNew.value = false;
        await load();
        toast('Gabarit créé');
      } catch(e) { toastErr(e); }
      finally { loading.value = false; }
    }

    async function openGabarit() {
      if (!openForm.path) return;
      loading.value = true;
      try {
        await POST('/api/gabarits/open', { ...openForm });
        Object.assign(openForm, { path: '', name: '', description: '' });
        showOpen.value = false;
        await load();
        toast('Gabarit référencé');
      } catch(e) { toastErr(e); }
      finally { loading.value = false; }
    }

    async function cloneGabarit() {
      if (!cloneForm.dest_path || !cloneForm.name) return;
      loading.value = true;
      try {
        await POST('/api/gabarits/clone', {
          source_path: showClone.value.path,
          dest_path: cloneForm.dest_path,
          name: cloneForm.name,
          activate: cloneForm.activate,
        });
        Object.assign(cloneForm, { dest_path: '', name: '', activate: true });
        showClone.value = null;
        await load();
        toast(cloneForm.activate ? 'Affaire créée et activée !' : 'Affaire créée');
      } catch(e) { toastErr(e); }
      finally { loading.value = false; }
    }

    async function exportCurrent() {
      if (!exportForm.name || !exportForm.dest_path) return;
      loading.value = true;
      try {
        await POST('/api/gabarits/from-current', { ...exportForm });
        Object.assign(exportForm, { name: '', dest_path: '', description: '' });
        showExport.value = false;
        await load();
        toast('Schéma N1 exporté comme Gabarit');
      } catch(e) { toastErr(e); }
      finally { loading.value = false; }
    }

    async function remove(name) {
      if (!confirm(`Retirer le Gabarit "${name}" de la liste ?`)) return;
      try {
        await DEL(`/api/gabarits/${encodeURIComponent(name)}`);
        await load();
        toast('Gabarit retiré');
      } catch(e) { toastErr(e); }
    }

    onMounted(load);
    return {
      gabarits, showNew, showOpen, showClone, showExport, loading,
      newForm, openForm, cloneForm, exportForm,
      createNew, openGabarit, cloneGabarit, exportCurrent, remove,
    };
  },
  template: `
    <div>
      <!-- Explication -->
      <div class="card" style="background:var(--surface2);border-left:3px solid var(--accent);">
        <div style="display:flex;gap:16px;align-items:flex-start;">
          <span style="font-size:1.4rem;">◈</span>
          <div>
            <div style="font-weight:600;margin-bottom:4px;">Gabarits & Affaires</div>
            <div style="font-size:0.8rem;color:var(--text-dim);line-height:1.5;">
              Un <strong>Gabarit</strong> est un schéma N1 pré-configuré (Classes, Relations, Fonctions).
              Il sert de point de départ réutilisable.<br>
              Une <strong>Affaire</strong> est un écosystème complet N1+N2 créé depuis un Gabarit.
              Elle vit de façon indépendante après le clonage.
            </div>
          </div>
        </div>
      </div>

      <!-- Liste des Gabarits -->
      <div class="card">
        <div class="card-header">
          <span class="card-title">Gabarits disponibles <span class="topbar-badge">{{ gabarits.length }}</span></span>
          <div style="display:flex;gap:8px;">
            <button class="btn btn-ghost btn-sm" @click="showExport=!showExport">↑ Exporter l'Affaire active</button>
            <button class="btn btn-ghost btn-sm" @click="showOpen=!showOpen">⊕ Référencer</button>
            <button class="btn btn-primary btn-sm" @click="showNew=!showNew">+ Nouveau</button>
          </div>
        </div>

        <!-- Formulaire : Nouveau Gabarit -->
        <div v-if="showNew" style="background:var(--surface2);border-radius:8px;padding:16px;margin-bottom:16px;">
          <div style="font-weight:600;margin-bottom:12px;">Nouveau Gabarit vide</div>
          <div class="form-row" style="margin-bottom:10px;">
            <div class="form-group" style="margin:0;">
              <label>Nom *</label>
              <input v-model="newForm.name" placeholder="Gabarit UO standard" />
            </div>
            <div class="form-group" style="margin:0;">
              <label>Chemin *</label>
              <input v-model="newForm.path" placeholder="gabarits/uo-standard" />
            </div>
          </div>
          <div class="form-group" style="margin-bottom:10px;">
            <label>Description</label>
            <input v-model="newForm.description" placeholder="Description optionnelle" />
          </div>
          <div style="display:flex;gap:8px;justify-content:flex-end;">
            <button class="btn btn-ghost btn-sm" @click="showNew=false">Annuler</button>
            <button class="btn btn-primary btn-sm" :disabled="loading" @click="createNew">Créer</button>
          </div>
        </div>

        <!-- Formulaire : Référencer un dossier existant -->
        <div v-if="showOpen" style="background:var(--surface2);border-radius:8px;padding:16px;margin-bottom:16px;">
          <div style="font-weight:600;margin-bottom:12px;">Référencer un dossier existant comme Gabarit</div>
          <div class="form-row" style="margin-bottom:10px;">
            <div class="form-group" style="margin:0;">
              <label>Chemin *</label>
              <input v-model="openForm.path" placeholder="chemin/vers/dossier-n1" />
            </div>
            <div class="form-group" style="margin:0;">
              <label>Nom (optionnel)</label>
              <input v-model="openForm.name" placeholder="Nom du gabarit" />
            </div>
          </div>
          <div style="display:flex;gap:8px;justify-content:flex-end;">
            <button class="btn btn-ghost btn-sm" @click="showOpen=false">Annuler</button>
            <button class="btn btn-primary btn-sm" :disabled="loading" @click="openGabarit">Référencer</button>
          </div>
        </div>

        <!-- Formulaire : Exporter l'Affaire active comme Gabarit -->
        <div v-if="showExport" style="background:var(--surface2);border-radius:8px;padding:16px;margin-bottom:16px;">
          <div style="font-weight:600;margin-bottom:12px;">Exporter le schéma N1 de l'Affaire active</div>
          <div class="form-row" style="margin-bottom:10px;">
            <div class="form-group" style="margin:0;">
              <label>Nom du Gabarit *</label>
              <input v-model="exportForm.name" placeholder="Gabarit MI20" />
            </div>
            <div class="form-group" style="margin:0;">
              <label>Destination *</label>
              <input v-model="exportForm.dest_path" placeholder="gabarits/mi20" />
            </div>
          </div>
          <div style="display:flex;gap:8px;justify-content:flex-end;">
            <button class="btn btn-ghost btn-sm" @click="showExport=false">Annuler</button>
            <button class="btn btn-primary btn-sm" :disabled="loading" @click="exportCurrent">Exporter</button>
          </div>
        </div>

        <!-- Tableau des Gabarits -->
        <table>
          <thead><tr>
            <th>Nom</th><th>Classes</th><th>Description</th><th>État</th><th></th>
          </tr></thead>
          <tbody>
            <tr v-if="gabarits.length === 0">
              <td colspan="5" style="text-align:center;color:var(--text-dim);padding:24px;">
                Aucun Gabarit — créez-en un ou référencez un dossier N1 existant.
              </td>
            </tr>
            <tr v-for="g in gabarits" :key="g.path">
              <td>
                <div style="font-weight:500;">{{ g.name }}</div>
                <div style="font-size:0.72rem;color:var(--text-dim);">{{ g.path }}</div>
              </td>
              <td><span class="badge badge-teal">{{ g.class_count }} classe(s)</span></td>
              <td style="font-size:0.8rem;color:var(--text-dim);">{{ g.description || '—' }}</td>
              <td><span class="badge" :class="g.valid ? 'badge-green' : 'badge-orange'">{{ g.valid ? 'valide' : 'invalide' }}</span></td>
              <td style="display:flex;gap:6px;justify-content:flex-end;">
                <button class="btn btn-primary btn-sm" @click="showClone = g">⎘ Cloner → Affaire</button>
                <button class="btn btn-danger btn-sm" @click="remove(g.name)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Modal : Cloner vers une Affaire -->
      <div v-if="showClone" class="modal-overlay" @click.self="showClone=null">
        <div class="modal">
          <div class="modal-title">Cloner "{{ showClone.name }}" → nouvelle Affaire</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px;">
            Le schéma N1 sera copié. L'Affaire sera ensuite indépendante du Gabarit.
          </div>
          <div class="form-group">
            <label>Nom de l'Affaire *</label>
            <input v-model="cloneForm.name" placeholder="Affaire Projet X" />
          </div>
          <div class="form-group">
            <label>Chemin de destination *</label>
            <input v-model="cloneForm.dest_path" placeholder="affaires/projet-x" />
          </div>
          <div class="form-group" style="display:flex;align-items:center;gap:8px;">
            <input type="checkbox" id="cb-activate" v-model="cloneForm.activate" />
            <label for="cb-activate" style="margin:0;cursor:pointer;">Activer immédiatement cette Affaire</label>
          </div>
          <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px;">
            <button class="btn btn-ghost" @click="showClone=null">Annuler</button>
            <button class="btn btn-primary" :disabled="loading" @click="cloneGabarit">⎘ Créer l'Affaire</button>
          </div>
        </div>
      </div>
    </div>`
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Workspace Global (partagé N1 + N2)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewWorkspaceGlobal = {
  setup() {
    const form    = ref({ gabarits_dir: '', workspace_dir: '' });
    const loading = ref(true);
    const saved   = ref(false);
    const err     = ref('');

    onMounted(async () => {
      try {
        const d = await GET('/api/workspace');
        form.value = { gabarits_dir: d.gabarits_dir || '', workspace_dir: d.workspace_dir || '' };
      } catch(e) { toastErr(e); }
      loading.value = false;
    });

    const save = async () => {
      err.value = '';
      try {
        await PUT('/api/workspace', form.value);
        saved.value = true;
        setTimeout(() => saved.value = false, 2000);
        toast('Workspace enregistré');
      } catch(e) { err.value = e.message; toastErr(e); }
    };

    const reset = async () => {
      if (!confirm('Réinitialiser le workspace global ?')) return;
      await DEL('/api/workspace');
      form.value = { gabarits_dir: '', workspace_dir: '' };
      toast('Workspace réinitialisé');
    };

    return { form, loading, saved, err, save, reset };
  },
  template: `
    <div v-if="loading" class="loading">Chargement…</div>
    <div v-else style="max-width:640px;">
      <div style="margin-bottom:20px;padding:14px 16px;background:var(--surface2);border-radius:8px;border-left:3px solid #a5b4fc;font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        Paramètres partagés entre N1 et N2. Stockés dans
        <code style="color:#a5b4fc;">.exosync_workspace.json</code> à la racine du projet.
      </div>
      <div class="form-group">
        <label class="form-label">Bibliothèque de Gabarits</label>
        <input v-model="form.gabarits_dir" class="form-control"
          placeholder="ex: C:\\ExoSync\\Gabarits" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Dossier scanné automatiquement pour lister les gabarits disponibles.
        </div>
      </div>
      <div class="form-group" style="margin-top:16px;">
        <label class="form-label">Répertoire Workspace (Affaires)</label>
        <input v-model="form.workspace_dir" class="form-control"
          placeholder="ex: C:\\ExoSync\\Affaires" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Dossier parent par défaut pour créer de nouvelles Affaires.
        </div>
      </div>
      <div v-if="err" style="margin-top:12px;padding:10px 14px;background:#7f1d1d;border-radius:6px;font-size:0.82rem;color:#fca5a5;">⚠ {{ err }}</div>
      <div style="display:flex;gap:10px;margin-top:20px;">
        <button class="btn btn-primary" @click="save">{{ saved ? '✓ Enregistré' : '💾 Enregistrer' }}</button>
        <button class="btn btn-ghost" @click="reset" style="color:#ef4444;border-color:#ef4444;">✕ Réinitialiser</button>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Répertoires de l'Affaire — Posts
// ═══════════════════════════════════════════════════════════════════════════════
const ViewDirectories = {
  setup() {
    const form    = ref({ posts_dir: '' });
    const saved   = ref(false);
    const loading = ref(true);
    const err     = ref('');

    onMounted(async () => {
      try {
        const d = await GET('/api/directories');
        form.value = { posts_dir: d.posts_dir || '' };
      } catch(e) { toastErr(e); }
      loading.value = false;
    });

    const save = async () => {
      err.value = '';
      try {
        await PUT('/api/directories', form.value);
        saved.value = true;
        setTimeout(() => saved.value = false, 2000);
        toast('Répertoire Posts enregistré');
      } catch(e) { err.value = e.message; toastErr(e); }
    };

    const reset = async () => {
      if (!confirm('Remettre le répertoire Posts à zéro ?')) return;
      await DEL('/api/directories');
      form.value = { posts_dir: '' };
      toast('Répertoire Posts réinitialisé');
    };

    return { form, saved, loading, err, save, reset };
  },
  template: `
    <div v-if="loading" class="loading">Chargement…</div>
    <div v-else style="max-width:640px;">

      <div style="margin-bottom:20px;padding:14px 16px;background:var(--surface2);border-radius:8px;border-left:3px solid var(--accent);font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        Définissez le dossier <strong style="color:var(--text)">Posts</strong> qui contient les fichiers Excel
        de cette Affaire. Ce chemin est utilisé par le moteur de synchronisation pour localiser les Posts.
      </div>

      <div class="form-group">
        <label class="form-label">Répertoire Posts (Excel)</label>
        <input v-model="form.posts_dir" class="form-control"
          placeholder="ex: C:\\\\Users\\\\user\\\\OneDrive\\\\ExoSync\\\\Posts" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Chemin <strong>absolu</strong>. Peut pointer vers OneDrive, un réseau ou tout autre emplacement.
          Un chemin non encore existant est accepté (synchronisation OneDrive différée).
        </div>
      </div>

      <div v-if="err" style="margin-top:12px;padding:10px 14px;background:#7f1d1d;border-radius:6px;font-size:0.82rem;color:#fca5a5;">
        ⚠ {{ err }}
      </div>

      <div style="display:flex;gap:10px;margin-top:20px;">
        <button class="btn btn-primary" @click="save">
          {{ saved ? '✓ Enregistré' : '💾 Enregistrer' }}
        </button>
        <button class="btn btn-ghost" @click="reset" style="color:#ef4444;border-color:#ef4444;">
          ✕ Réinitialiser
        </button>
      </div>

      <div style="margin-top:24px;padding:12px 16px;background:var(--surface2);border-radius:8px;font-size:0.78rem;color:var(--text-dim);">
        <div style="font-weight:600;color:var(--text);margin-bottom:6px;">Impact sur le moteur Sync</div>
        <div>Les chemins relatifs des Posts dans le Registre sont résolus comme :</div>
        <div style="font-family:monospace;margin:6px 0;color:var(--accent);">
          {{ form.posts_dir || '[posts_dir]' }} \\ [chemin du Post]
        </div>
        <div>Les chemins absolus dans le Registre ne sont pas affectés.</div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Rôles & Fonctions (CRUD — source de vérité partagée avec N1)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewFunctions = {
  setup() {
    const items     = ref([]);
    const showModal = ref(false);
    const editing   = ref(null);

    const blank = () => ({ id:'', label:'', side:'interne', description:'' });
    const form  = reactive(blank());

    const load = async () => {
      try { items.value = await GET('/api/functions'); } catch(e) { toastErr(e); }
    };
    onMounted(load);

    const openCreate = () => { Object.assign(form, blank()); editing.value = null; showModal.value = true; };
    const openEdit   = item => { Object.assign(form, { ...item }); editing.value = item.id; showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value) { await PUT(`/api/functions/${editing.value}`, form); toast('Rôle mis à jour'); }
        else               { await POST('/api/functions', form);              toast('Rôle créé'); }
        showModal.value = false; await load();
      } catch(e) { toastErr(e); }
    };

    const del = async id => {
      if (!confirm(`Supprimer le rôle "${id}" ? Les acteurs associés devront être mis à jour.`)) return;
      try { await DEL(`/api/functions/${id}`); await load(); toast('Rôle supprimé'); }
      catch(e) { toastErr(e); }
    };

    return { items, showModal, editing, form, openCreate, openEdit, save, del };
  },
  template: `
    <div>
      <!-- Info -->
      <div style="margin-bottom:16px;padding:12px 16px;background:var(--surface2);border-radius:8px;
                  border-left:3px solid #a5b4fc;font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        Les <strong style="color:var(--text)">Rôles & Fonctions</strong> sont stockés dans
        <code style="color:#a5b4fc;">functions.json</code> et partagés avec N1 (Schema Designer).
        Toute modification ici est immédiatement reflétée dans les deux applications.
      </div>

      <div class="card">
        <div class="card-header">
          <span class="card-title">
            Rôles
            <span class="topbar-badge">{{ items.length }}</span>
            <span style="margin-left:8px;font-size:0.75rem;color:var(--text-dim)">
              — {{ items.filter(f=>f.side==='interne').length }} internes ·
                 {{ items.filter(f=>f.side==='externe').length }} externes
            </span>
          </span>
          <button class="btn btn-primary" @click="openCreate">+ Nouveau rôle</button>
        </div>
        <table>
          <thead>
            <tr>
              <th>ID technique</th>
              <th>Libellé</th>
              <th>Côté</th>
              <th>Description</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="f in items" :key="f.id">
              <td><code style="color:var(--accent);font-size:0.8rem">{{ f.id }}</code></td>
              <td><strong>{{ f.label }}</strong></td>
              <td>
                <span class="badge" :class="f.side==='externe' ? 'badge-green' : 'badge-blue'">
                  {{ f.side }}
                </span>
              </td>
              <td style="font-size:0.78rem;color:var(--text-dim)">{{ f.description }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(f)">Éditer</button>
                <button class="btn btn-danger btn-sm" @click="del(f.id)">✕</button>
              </td>
            </tr>
            <tr v-if="!items.length">
              <td colspan="5" style="text-align:center;color:var(--text-dim);padding:20px">
                Aucun rôle défini — cliquez sur "+ Nouveau rôle"
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Modal CRUD -->
      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier le rôle' : 'Nouveau rôle' }}</div>

          <div class="form-row">
            <div class="form-group">
              <label>ID technique *</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="ex: chef_projet" />
              <div style="font-size:0.75rem;color:var(--text-dim);margin-top:3px">
                Identifiant interne (snake_case, sans espaces)
              </div>
            </div>
            <div class="form-group">
              <label>Libellé *</label>
              <input v-model="form.label" placeholder="ex: Chef de Projet" />
            </div>
          </div>

          <div class="form-group" style="margin-top:12px">
            <label>Côté</label>
            <select v-model="form.side">
              <option value="interne">Interne (équipe projet)</option>
              <option value="externe">Externe (client, fournisseur)</option>
            </select>
          </div>

          <div class="form-group" style="margin-top:12px">
            <label>Description</label>
            <textarea v-model="form.description" rows="2"
              placeholder="Responsabilités, périmètre d'action…"
              style="width:100%;resize:vertical;"></textarea>
          </div>

          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">{{ editing ? 'Enregistrer' : 'Créer' }}</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// ROOT APP
// ═══════════════════════════════════════════════════════════════════════════════
const VIEW_MAP = {
  dashboard:      ViewDashboard,
  workspace:      ViewEcosystemManager,
  gabarits:       ViewGabarits,
  wsglobal:       ViewWorkspaceGlobal,
  directories:    ViewDirectories,
  sync:           ViewSync,
  registry:       ViewRegistry,
  actors:         ViewActors,
  functions:      ViewFunctions,
  hierarchy:      ViewHierarchy,
  tables:         ViewTables,
  tissage:        ViewTissage,
  mxl:            ViewMxl,
  'excel-import': ViewExcelImport,
  ecosystem:      ViewEcosystem,
};

const App = {
  components: { ToastLayer },
  setup() {
    const view = ref('dashboard');
    const ecosystemStatus = ref(null);
    const ecosystemCurrent = ref(null);

    onMounted(async () => {
      [ecosystemStatus.value, ecosystemCurrent.value] = await Promise.all([
        GET('/api/ecosystem/status').catch(() => null),
        GET('/api/ecosystem/current').catch(() => null),
      ]);
    });

    const nav = [
      { id:'dashboard',    icon:'⊞', label:'Tableau de bord',        group:'Principal' },
      { id:'workspace',    icon:'◎', label:'Espace de travail',      group:'Configuration' },
      { id:'gabarits',     icon:'◈', label:'Gabarits & Affaires',    group:'Configuration' },
      { id:'wsglobal',     icon:'⌂', label:'Workspace global',       group:'Configuration' },
      { id:'registry',     icon:'◧', label:'Registre des Posts',     group:'Population' },
      { id:'actors',       icon:'◉', label:'Acteurs',              group:'Population' },
      { id:'functions',    icon:'◈', label:'Rôles & Fonctions',    group:'Population' },
      { id:'excel-import', icon:'⤵', label:'Importer Excel',       group:'Population' },
      { id:'hierarchy',    icon:'⬡', label:'Hiérarchie LIST+COLLECT', group:'Structure' },
      { id:'tables',       icon:'▦', label:'Tables & colonnes',  group:'Structure' },
      { id:'directories',  icon:'⌂', label:'Répertoires',          group:'Configuration' },
      { id:'sync',         icon:'↻', label:'Synchronisation',     group:'Flux' },
      { id:'tissage',      icon:'⇢', label:'Tissage PULL',       group:'Flux' },
      { id:'mxl',          icon:'⌥', label:'Générateur MXL',     group:'Flux' },
      { id:'ecosystem',    icon:'◎', label:'Graphe écosystème',  group:'Visualisation' },
    ];

    const navGroups = computed(() => {
      const g = {};
      nav.forEach(n => { if (!g[n.group]) g[n.group] = []; g[n.group].push(n); });
      return g;
    });

    const titles = {
      dashboard:    'Tableau de bord',
      workspace:    'Gestion des espaces de travail',
      gabarits:     'Gabarits & Affaires',
      registry:     'Registre des Posts',
      actors:       'Acteurs',
      functions:    'Rôles & Fonctions',
      hierarchy:    'Hiérarchie — LIST & COLLECT',
      tables:       'Tables & colonnes',
      directories:  'Répertoires — Affaire & Posts',
      wsglobal:     'Workspace Global — Gabarits & Affaires',
      sync:         'Synchronisation ExoSync',
      tissage:      'Tissage — Déclaration de flux PULL',
      mxl:          'Générateur de Manifeste MXL',
      'excel-import': 'Importer un fichier Excel',
      ecosystem:    'Graphe de l\'écosystème',
    };

    const currentView = computed(() => VIEW_MAP[view.value] || ViewDashboard);

    return { view, navGroups, titles, ecosystemStatus, ecosystemCurrent, toasts, currentView };
  },
  template: `
    <div id="sidebar">
      <div class="sidebar-logo">
        ExoSync<br>
        <span>Studio N2 · Registry Populator</span>
      </div>
      <nav>
        <div v-for="(items, group) in navGroups" :key="group" class="nav-group">
          <div class="nav-group-label">{{ group }}</div>
          <div v-for="item in items" :key="item.id"
               class="nav-item" :class="{ active: view === item.id }"
               @click="view = item.id">
            <span class="nav-icon">{{ item.icon }}</span>
            {{ item.label }}
          </div>
        </div>
      </nav>
      <div class="sidebar-status">
        <span class="status-dot" :class="ecosystemStatus?.errors===0?'ok':'err'"></span>
        <span v-if="ecosystemCurrent" style="font-weight:600;display:block;margin-bottom:2px;">{{ ecosystemCurrent.name }}</span>
        {{ ecosystemStatus ? ecosystemStatus.total_files + ' Posts · ' + ecosystemStatus.edge_count + ' arêtes' : 'Chargement…' }}
      </div>
    </div>

    <div id="main">
      <div class="topbar">
        <span class="topbar-title">{{ titles[view] }}</span>
        <span class="topbar-spacer"></span>
        <span style="font-size:0.72rem;color:var(--text-dim)">N2 — Population</span>
      </div>
      <div class="content">
        <component :is="currentView" :key="view" />
      </div>
    </div>

    <toast-layer :toasts="toasts" />
  `
};

createApp(App).mount('#app');
