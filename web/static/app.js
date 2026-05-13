const { createApp, ref, reactive, computed, onMounted, watch, nextTick } = Vue;

// ── API helpers ───────────────────────────────────────────────────────────────
const BASE = '';
async function api(method, path, body) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (body !== undefined) opts.body = JSON.stringify(body);
  const r = await fetch(BASE + path, opts);
  if (!r.ok) {
    const err = await r.json().catch(() => ({ detail: r.statusText }));
    throw new Error(err.detail || r.statusText);
  }
  if (r.status === 204) return null;
  return r.json();
}
const GET = path => api('GET', path);
const POST = (path, body) => api('POST', path, body);
const PUT = (path, body) => api('PUT', path, body);
const DEL = path => api('DELETE', path);

// ── Role colors — alignés sur SCHEMA-01 ──────────────────────────────────────
const ROLE_BADGE = {
  // Internes
  pilote_pole:          'badge-orange',
  pilote_metier:        'badge-green',
  ingenieur_sys:        'badge-blue',
  it_manager:           'badge-purple',
  // Externes — donneurs d'ordre
  donneur_ordre_pole:   'badge-orange',
  donneur_ordre_metier: 'badge-gray',
  // Externes — correspondants techniques
  ingenieur_sys_client: 'badge-green',
  expert_technique:     'badge-purple',
  fournisseur:          'badge-gray',
  correspondant_projet: 'badge-gray',
};

const FILE_TYPE_BADGE = {
  uo_instance: 'badge-blue',
  referentiel_uo: 'badge-purple',
  referentiel_projet: 'badge-orange',
  cockpit: 'badge-green',
  pilote: 'badge-green',
  consolidation: 'badge-orange',
  client: 'badge-gray',
};

// ── Toast ─────────────────────────────────────────────────────────────────────
const ToastComponent = {
  props: ['toasts'],
  template: `
    <div style="position:fixed;bottom:24px;right:24px;z-index:999;display:flex;flex-direction:column;gap:8px;">
      <transition-group name="toast">
        <div v-for="t in toasts" :key="t.id"
          :style="{background: t.type==='error' ? '#7f1d1d' : '#14532d',
                   border: '1px solid ' + (t.type==='error' ? '#ef4444' : '#22c55e'),
                   padding:'10px 16px', borderRadius:'8px', fontSize:'0.85rem',
                   color:'#fff', maxWidth:'320px', boxShadow:'0 4px 12px rgba(0,0,0,0.4)'}">
          {{ t.msg }}
        </div>
      </transition-group>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Dashboard
// ═══════════════════════════════════════════════════════════════════════════════
const ViewDashboard = {
  setup() {
    const status = ref(null);
    const fileTypes = ref([]);
    const actors = ref([]);
    const registry = ref([]);

    onMounted(async () => {
      [status.value, fileTypes.value, actors.value, registry.value] = await Promise.all([
        GET('/api/ecosystem/status'),
        GET('/api/file-types'),
        GET('/api/actors'),
        GET('/api/registry'),
      ]);
    });

    const syncOk = computed(() => registry.value.filter(f => f.statut_dernier_synchro === 'ok').length);
    const syncErr = computed(() => registry.value.filter(f => f.statut_dernier_synchro && f.statut_dernier_synchro !== 'ok').length);

    return { status, fileTypes, actors, registry, syncOk, syncErr };
  },
  template: `
    <div>
      <div class="stats-row">
        <div class="stat-card">
          <div class="stat-value" style="color:#3b82f6">{{ fileTypes.length }}</div>
          <div class="stat-label">Types de fichiers</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#8b5cf6">{{ registry.length }}</div>
          <div class="stat-label">Fichiers enregistrés</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#10b981">{{ actors.length }}</div>
          <div class="stat-label">Acteurs</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#22c55e">{{ syncOk }}</div>
          <div class="stat-label">Syncs OK</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color:#f59e0b">{{ status ? status.edge_count : '…' }}</div>
          <div class="stat-label">Arêtes écosystème</div>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
        <div class="card">
          <div class="card-header"><span class="card-title">Fichiers récents</span></div>
          <table>
            <thead><tr><th>ID</th><th>Type</th><th>Statut</th><th>Dernière sync</th></tr></thead>
            <tbody>
              <tr v-for="f in registry.slice(0,6)" :key="f.id">
                <td><strong>{{ f.id }}</strong></td>
                <td><span class="badge" :class="FILE_TYPE_BADGE[f.type_fichier]||'badge-gray'">{{ f.type_fichier }}</span></td>
                <td>
                  <span class="badge" :class="f.statut_dernier_synchro==='ok'?'badge-green':'badge-orange'">
                    {{ f.statut_dernier_synchro || 'N/A' }}
                  </span>
                </td>
                <td style="color:var(--text-dim);font-size:0.78rem">{{ f.derniere_synchro ? f.derniere_synchro.slice(0,16) : '—' }}</td>
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
  `,
  computed: {
    FILE_TYPE_BADGE() { return FILE_TYPE_BADGE; },
    ROLE_BADGE() { return ROLE_BADGE; },
    roleGroups() {
      const m = {};
      this.actors.forEach(a => { m[a.role] = (m[a.role] || 0) + 1; });
      return Object.entries(m);
    }
  }
};

// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: FileTypes
// ═══════════════════════════════════════════════════════════════════════════════
const ViewFileTypes = {
  components: { TagsInput },
  emits: ['toast'],
  setup(_, { emit }) {
    const items = ref([]);
    const showModal = ref(false);
    const editing = ref(null);
    const ROLES = ['pilote_pole','pilote_metier','ingenieur_sys','it_manager','donneur_ordre_pole','donneur_ordre_metier','ingenieur_sys_client','expert_technique','fournisseur','correspondant_projet'];

    const blank = () => ({
      id: '', label: '', description: '', template: '', owner_role: 'ingenieur_sys',
      required_sheets: [], optional_sheets: [], allowed_namespaces: [], push_prefix: ''
    });
    const form = reactive(blank());

    const load = async () => { items.value = await GET('/api/file-types'); };
    onMounted(load);

    const openCreate = () => { Object.assign(form, blank()); editing.value = null; showModal.value = true; };
    const openEdit = (item) => {
      Object.assign(form, { ...item });
      editing.value = item.id;
      showModal.value = true;
    };

    const save = async () => {
      try {
        if (editing.value) {
          await PUT(`/api/file-types/${editing.value}`, form);
          emit('toast', { msg: 'Type mis à jour', type: 'ok' });
        } else {
          await POST('/api/file-types', form);
          emit('toast', { msg: 'Type créé', type: 'ok' });
        }
        showModal.value = false;
        await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const del = async (id) => {
      if (!confirm(`Supprimer le type "${id}" ?`)) return;
      try { await DEL(`/api/file-types/${id}`); await load(); emit('toast', { msg: 'Type supprimé', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    return { items, showModal, editing, form, ROLES, openCreate, openEdit, save, del, FILE_TYPE_BADGE };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <span class="card-title">Types de fichiers <span class="topbar-badge">{{ items.length }}</span></span>
          <button class="btn btn-primary" @click="openCreate">+ Nouveau type</button>
        </div>
        <table>
          <thead><tr><th>ID</th><th>Label</th><th>Rôle owner</th><th>Feuilles requises</th><th>Namespace</th><th></th></tr></thead>
          <tbody>
            <tr v-for="t in items" :key="t.id">
              <td><code style="color:var(--accent)">{{ t.id }}</code></td>
              <td>{{ t.label }}</td>
              <td><span class="badge" :class="'badge-blue'">{{ t.owner_role }}</span></td>
              <td style="font-size:0.78rem">{{ (t.required_sheets||[]).join(', ') }}</td>
              <td style="font-size:0.78rem;color:var(--text-dim)">{{ (t.allowed_namespaces||[]).join(' ') }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(t)">Éditer</button>
                <button class="btn btn-danger btn-sm" @click="del(t.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouveau' }} type de fichier</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID *</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="uo_instance" />
            </div>
            <div class="form-group">
              <label>Label *</label>
              <input v-model="form.label" placeholder="Unite d Oeuvre" />
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="form.description" rows="2"></textarea>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Rôle owner *</label>
              <select v-model="form.owner_role">
                <option v-for="r in ROLES" :key="r" :value="r">{{ r }}</option>
              </select>
            </div>
            <div class="form-group">
              <label>Push prefix</label>
              <input v-model="form.push_prefix" placeholder="uo.{id}." />
            </div>
          </div>
          <div class="form-group">
            <label>Template MXL (chemin)</label>
            <input v-model="form.template" placeholder="config/templates/manifeste_uo_instance.mxl" />
          </div>
          <div class="form-group">
            <label>Feuilles requises</label>
            <tags-input v-model="form.required_sheets" placeholder="_Manifeste, Activites…" />
          </div>
          <div class="form-group">
            <label>Feuilles optionnelles</label>
            <tags-input v-model="form.optional_sheets" placeholder="Dashboard, _Log…" />
          </div>
          <div class="form-group">
            <label>Namespaces autorisés</label>
            <tags-input v-model="form.allowed_namespaces" placeholder="uo. ref. projet.…" />
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
// VIEW: Registry
// ═══════════════════════════════════════════════════════════════════════════════
const ViewRegistry = {
  emits: ['toast'],
  setup(_, { emit }) {
    const items = ref([]);
    const fileTypes = ref([]);
    const actors = ref([]);
    const showModal = ref(false);
    const editing = ref(null);
    const PERIODICITES = ['quotidien','hebdomadaire','manuel'];

    const blank = () => ({
      id: '', type_fichier: '', chemin: '', synchro_periodicite: 'quotidien',
      owner_id: '', genere_par_script: false
    });
    const form = reactive(blank());

    const load = async () => {
      [items.value, fileTypes.value, actors.value] = await Promise.all([
        GET('/api/registry'), GET('/api/file-types'), GET('/api/actors')
      ]);
    };
    onMounted(load);

    const openCreate = () => { Object.assign(form, blank()); editing.value = null; showModal.value = true; };
    const openEdit = item => { Object.assign(form, { ...item }); editing.value = item.id; showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value) { await PUT(`/api/registry/${editing.value}`, form); emit('toast', { msg: 'Fichier mis à jour', type: 'ok' }); }
        else { await POST('/api/registry', form); emit('toast', { msg: 'Fichier ajouté', type: 'ok' }); }
        showModal.value = false; await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const del = async id => {
      if (!confirm(`Supprimer "${id}" du registre ?`)) return;
      try { await DEL(`/api/registry/${id}`); await load(); emit('toast', { msg: 'Fichier supprimé', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const ownerName = id => actors.value.find(a => a.id === id)?.nom || id;

    return { items, fileTypes, actors, showModal, editing, form, PERIODICITES,
             openCreate, openEdit, save, del, ownerName, FILE_TYPE_BADGE };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <span class="card-title">Registre des fichiers <span class="topbar-badge">{{ items.length }}</span></span>
          <button class="btn btn-primary" @click="openCreate">+ Ajouter un fichier</button>
        </div>
        <table>
          <thead><tr><th>ID</th><th>Type</th><th>Owner</th><th>Périodicité</th><th>Statut</th><th>Chemin</th><th></th></tr></thead>
          <tbody>
            <tr v-for="f in items" :key="f.id">
              <td><strong>{{ f.id }}</strong></td>
              <td><span class="badge" :class="FILE_TYPE_BADGE[f.type_fichier]||'badge-gray'">{{ f.type_fichier }}</span></td>
              <td style="font-size:0.82rem">{{ ownerName(f.owner_id) }}</td>
              <td><span class="badge badge-gray">{{ f.synchro_periodicite }}</span></td>
              <td>
                <span class="badge" :class="f.statut_dernier_synchro==='ok'?'badge-green':'badge-orange'">
                  {{ f.statut_dernier_synchro || 'N/A' }}
                </span>
              </td>
              <td style="font-size:0.75rem;color:var(--text-dim);max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">{{ f.chemin }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <a :href="'/api/xlsx/' + f.id" target="_blank" class="btn btn-ghost btn-sm" title="Télécharger le fichier Excel squelette">⬇ XLS</a>
                <button class="btn btn-ghost btn-sm" @click="openEdit(f)">Éditer</button>
                <button class="btn btn-danger btn-sm" @click="del(f.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Ajouter' }} un fichier</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID *</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="UO-006" />
            </div>
            <div class="form-group">
              <label>Type *</label>
              <select v-model="form.type_fichier">
                <option value="">— Choisir —</option>
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
              <label>Owner *</label>
              <select v-model="form.owner_id">
                <option value="">— Choisir —</option>
                <option v-for="a in actors" :key="a.id" :value="a.id">{{ a.nom }} ({{ a.role }})</option>
              </select>
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
// VIEW: Actors
// ═══════════════════════════════════════════════════════════════════════════════
const ViewActors = {
  emits: ['toast'],
  setup(_, { emit }) {
    const items = ref([]);
    const showModal = ref(false);
    const editing = ref(null);
    const ROLES = ['pilote_pole','pilote_metier','ingenieur_sys','it_manager','donneur_ordre_pole','donneur_ordre_metier','ingenieur_sys_client','expert_technique','fournisseur','correspondant_projet'];
    const ACCES = ['read','read/write','read_filtered','read_summary','admin'];
    const FILTRE_TYPES = ['ingenieur','projet','ALL'];

    const blank = () => ({ id:'', nom:'', role:'ingenieur_sys', filtre_type:'ALL', filtre_valeur:'ALL', acces:'read', email:'' });
    const form = reactive(blank());

    const load = async () => { items.value = await GET('/api/actors'); };
    onMounted(load);

    const openCreate = () => { Object.assign(form, blank()); editing.value = null; showModal.value = true; };
    const openEdit = item => { Object.assign(form, { ...item }); editing.value = item.id; showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value) { await PUT(`/api/actors/${editing.value}`, form); emit('toast', { msg: 'Acteur mis à jour', type: 'ok' }); }
        else { await POST('/api/actors', form); emit('toast', { msg: 'Acteur créé', type: 'ok' }); }
        showModal.value = false; await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const del = async id => {
      if (!confirm(`Supprimer l'acteur "${id}" ?`)) return;
      try { await DEL(`/api/actors/${id}`); await load(); emit('toast', { msg: 'Acteur supprimé', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    return { items, showModal, editing, form, ROLES, ACCES, FILTRE_TYPES, openCreate, openEdit, save, del, ROLE_BADGE };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <span class="card-title">Acteurs <span class="topbar-badge">{{ items.length }}</span></span>
          <button class="btn btn-primary" @click="openCreate">+ Nouvel acteur</button>
        </div>
        <table>
          <thead><tr><th>ID</th><th>Nom</th><th>Rôle</th><th>Accès</th><th>Filtre</th><th>Email</th><th></th></tr></thead>
          <tbody>
            <tr v-for="a in items" :key="a.id">
              <td><code style="color:var(--accent)">{{ a.id }}</code></td>
              <td><strong>{{ a.nom }}</strong></td>
              <td><span class="badge" :class="ROLE_BADGE[a.role]||'badge-gray'">{{ a.role }}</span></td>
              <td><span class="badge badge-gray">{{ a.acces }}</span></td>
              <td style="font-size:0.78rem;color:var(--text-dim)">{{ a.filtre_type }} : {{ a.filtre_valeur }}</td>
              <td style="font-size:0.8rem;color:var(--text-dim)">{{ a.email }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(a)">Éditer</button>
                <button class="btn btn-danger btn-sm" @click="del(a.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvel' }} acteur</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID *</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="USR009" />
            </div>
            <div class="form-group">
              <label>Nom *</label>
              <input v-model="form.nom" placeholder="Prénom Nom" />
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Rôle *</label>
              <select v-model="form.role">
                <option v-for="r in ROLES" :key="r" :value="r">{{ r }}</option>
              </select>
            </div>
            <div class="form-group">
              <label>Accès</label>
              <select v-model="form.acces">
                <option v-for="a in ACCES" :key="a" :value="a">{{ a }}</option>
              </select>
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Type de filtre</label>
              <select v-model="form.filtre_type">
                <option v-for="f in FILTRE_TYPES" :key="f" :value="f">{{ f }}</option>
              </select>
            </div>
            <div class="form-group">
              <label>Valeur du filtre</label>
              <input v-model="form.filtre_valeur" placeholder="ALL ou nom/projet" />
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
// VIEW: Ecosystem Graph
// ═══════════════════════════════════════════════════════════════════════════════
const NODE_COLORS = {
  uo_instance: '#3b82f6', referentiel_uo: '#8b5cf6', referentiel_projet: '#ec4899',
  cockpit: '#f59e0b', pilote: '#10b981', consolidation: '#f97316', client: '#06b6d4',
  store: '#475569', unknown: '#64748b',
};

const ViewEcosystem = {
  setup() {
    const graphData = ref(null);
    const status = ref(null);
    const cyInstance = ref(null);
    const detail = ref(null);
    const filter = ref('');

    const load = async () => {
      [graphData.value, status.value] = await Promise.all([
        GET('/api/ecosystem/graph'), GET('/api/ecosystem/status')
      ]);
    };

    onMounted(async () => {
      await load();
      await nextTick();
      buildGraph();
    });

    const buildGraph = () => {
      if (!graphData.value || typeof cytoscape === 'undefined') return;

      const fileNodes = graphData.value.nodes;
      const edges = graphData.value.edges;

      // Collect store nodes from edges
      const storeIds = new Set();
      edges.forEach(e => {
        [e.from_node, e.to_node].forEach(n => {
          if (n.startsWith('store::')) storeIds.add(n.split('::')[0] + '::' + n.split('::')[1]);
        });
      });

      const elements = [];

      // File nodes
      fileNodes.forEach(n => {
        elements.push({
          data: { id: n.id, label: n.id, type: n.file_type, status: n.status, path: n.path,
                  color: NODE_COLORS[n.file_type] || NODE_COLORS.unknown }
        });
      });

      // Store pseudo-nodes (aggregate by store key prefix)
      const storeGroups = {};
      edges.forEach(e => {
        [e.from_node, e.to_node].forEach(n => {
          if (n.startsWith('store::')) {
            const key = n.replace('store::', '').split('.')[0];
            storeGroups[key] = true;
          }
        });
      });
      Object.keys(storeGroups).forEach(k => {
        const sid = 'store_' + k;
        elements.push({ data: { id: sid, label: 'store::\n' + k + '.*', type: 'store', color: NODE_COLORS.store } });
      });

      // Edges — group file→store and store→file
      const edgeMap = {};
      edges.forEach(e => {
        const fromNode = resolveNodeId(e.from_node);
        const toNode   = resolveNodeId(e.to_node);
        if (!fromNode || !toNode || fromNode === toNode) return;
        const key = `${fromNode}→${toNode}→${e.edge_type}`;
        edgeMap[key] = { from: fromNode, to: toNode, type: e.edge_type };
      });

      Object.entries(edgeMap).forEach(([k, e]) => {
        elements.push({
          data: { id: 'e_'+k, source: e.from, target: e.to, edgeType: e.type,
                  color: e.type === 'PUSH' ? '#22c55e' : e.type === 'PULL' ? '#60a5fa' : '#f97316' }
        });
      });

      const cy = cytoscape({
        container: document.getElementById('cy'),
        elements,
        style: [
          {
            selector: 'node',
            style: {
              'background-color': 'data(color)',
              'label': 'data(label)',
              'color': '#e2e8f0',
              'font-size': '11px',
              'text-valign': 'bottom',
              'text-halign': 'center',
              'text-margin-y': '4px',
              'width': 48, 'height': 48,
              'border-width': 2,
              'border-color': '#1e293b',
              'text-wrap': 'wrap',
              'text-max-width': '80px',
            }
          },
          {
            selector: 'node[type="store"]',
            style: { shape: 'rectangle', width: 60, height: 36, 'font-size': '10px' }
          },
          {
            selector: 'edge',
            style: {
              'line-color': 'data(color)',
              'target-arrow-color': 'data(color)',
              'target-arrow-shape': 'triangle',
              'curve-style': 'bezier',
              'width': 2,
              'label': 'data(edgeType)',
              'font-size': '9px',
              'color': '#94a3b8',
              'text-rotation': 'autorotate',
            }
          },
          { selector: ':selected', style: { 'border-width': 3, 'border-color': '#fff' } },
          { selector: '.faded', style: { opacity: 0.15 } },
        ],
        layout: { name: 'cose', idealEdgeLength: 140, nodeOverlap: 20, refresh: 20,
                  fit: true, padding: 40, randomize: false, animate: false }
      });

      cy.on('tap', 'node', evt => {
        const n = evt.target.data();
        detail.value = n;
        cy.elements().addClass('faded');
        evt.target.removeClass('faded');
        evt.target.neighborhood().removeClass('faded');
      });
      cy.on('tap', evt => {
        if (evt.target === cy) { detail.value = null; cy.elements().removeClass('faded'); }
      });

      cyInstance.value = cy;
    };

    const resolveNodeId = (node) => {
      if (!node) return null;
      if (node.startsWith('store::')) {
        const key = node.replace('store::', '').split('.')[0];
        return 'store_' + key;
      }
      // node like "UO-001::TabActivites" → just file id
      return node.split('::')[0];
    };

    const relayout = () => { if (cyInstance.value) cyInstance.value.layout({ name: 'cose', animate: false, fit: true, padding: 40 }).run(); };
    const fitGraph = () => { if (cyInstance.value) cyInstance.value.fit(40); };

    return { graphData, status, detail, filter, relayout, fitGraph };
  },
  template: `
    <div>
      <div v-if="status" style="display:flex;gap:12px;margin-bottom:16px;align-items:center">
        <div class="stat-card" style="padding:12px 20px;display:flex;align-items:center;gap:12px">
          <div style="font-size:1.4rem;font-weight:700;color:#22c55e">{{ status.ok }}</div>
          <div style="font-size:0.75rem;color:var(--text-dim)">fichiers OK</div>
        </div>
        <div class="stat-card" style="padding:12px 20px;display:flex;align-items:center;gap:12px">
          <div style="font-size:1.4rem;font-weight:700;color:#f59e0b">{{ status.edge_count }}</div>
          <div style="font-size:0.75rem;color:var(--text-dim)">arêtes</div>
        </div>
        <div style="font-size:0.78rem;color:var(--text-dim)">Dernier scan : {{ status.last_scan ? status.last_scan.slice(0,16) : '—' }}</div>
        <div style="flex:1"></div>
        <button class="btn btn-ghost btn-sm" @click="fitGraph">⊡ Recadrer</button>
        <button class="btn btn-ghost btn-sm" @click="relayout">↺ Relayout</button>
      </div>

      <div style="position:relative">
        <div id="cy"></div>
        <div v-if="detail" style="position:absolute;top:12px;right:12px;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px;min-width:200px;font-size:0.82rem">
          <div style="font-weight:600;margin-bottom:8px">{{ detail.label }}</div>
          <div v-if="detail.type !== 'store'">
            <div style="color:var(--text-dim)">Type</div>
            <div class="badge badge-blue" style="margin-bottom:8px">{{ detail.type }}</div><br>
            <div style="color:var(--text-dim)">Statut</div>
            <div class="badge" :class="detail.status==='ok'?'badge-green':'badge-orange'">{{ detail.status }}</div><br><br>
            <div style="color:var(--text-dim);font-size:0.75rem;word-break:break-all">{{ detail.path }}</div>
          </div>
          <div v-else style="color:var(--text-dim)">Nœud store central</div>
        </div>
      </div>

      <div style="display:flex;gap:12px;margin-top:12px;font-size:0.78rem;color:var(--text-dim);flex-wrap:wrap">
        <span style="display:flex;align-items:center;gap:4px"><span style="width:24px;height:3px;background:#22c55e;display:inline-block"></span> PUSH</span>
        <span style="display:flex;align-items:center;gap:4px"><span style="width:24px;height:3px;background:#60a5fa;display:inline-block"></span> PULL</span>
        <span style="display:flex;align-items:center;gap:4px"><span style="width:24px;height:3px;background:#f97316;display:inline-block"></span> COLLECT</span>
        <span style="margin-left:16px">Cliquer un nœud pour voir les connexions</span>
      </div>
    </div>
  `
};

// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: ExcelImport (wizard 5 étapes)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewExcelImport = {
  emits: ['toast'],
  setup(_, { emit }) {
    const step = ref(1);
    const scanning = ref(false);
    const saving = ref(false);
    const comparing = ref(false);
    const dragOver = ref(false);
    const uploadedFile = ref(null);
    const scanResult = ref(null);
    const fileTypes = ref([]);
    const actors = ref([]);
    const availableTables = ref({});
    const selectedClass = ref('');
    const tableComparisons = ref({});
    const tableMappings = ref([]);
    const collectMappings = ref([]);
    const form = reactive({ file_id: '', owner_role: '', periodicite: 'quotidien', update_class_fields: false });

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
      [fileTypes.value, actors.value] = await Promise.all([GET('/api/file-types'), GET('/api/actors')]);
    });

    const onDrop = e => {
      dragOver.value = false;
      const file = e.dataTransfer?.files?.[0];
      if (file) handleFile(file);
    };

    const onFileInput = e => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
    };

    const handleFile = file => {
      if (!file.name.match(/\.(xlsx|xlsm|xlsb|xls)$/i)) {
        emit('toast', { msg: 'Le fichier doit être un .xlsx / .xlsm', type: 'error' });
        return;
      }
      uploadedFile.value = file;
    };

    const scanFile = async () => {
      if (!uploadedFile.value) return;
      scanning.value = true;
      try {
        const fd = new FormData();
        fd.append('file', uploadedFile.value);
        const r = await fetch('/api/excel/scan-upload', { method: 'POST', body: fd });
        if (!r.ok) {
          const err = await r.json().catch(() => ({ detail: r.statusText }));
          throw new Error(err.detail || r.statusText);
        }
        scanResult.value = await r.json();
        const idRange = scanResult.value.named_ranges.find(nr =>
          ['uo_id', 'file_id', 'post_id', 'id'].includes(nr.name.toLowerCase())
        );
        if (idRange?.value_preview) form.file_id = idRange.value_preview;
        step.value = 2;
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
      finally { scanning.value = false; }
    };

    const compareAndProceed = async () => {
      if (!selectedClass.value) { emit('toast', { msg: 'Choisir une Classe', type: 'error' }); return; }
      comparing.value = true;
      try {
        if (Object.keys(availableTables.value).length === 0) {
          const data = await GET('/api/excel/available-tables');
          availableTables.value = data.by_class || {};
        }
        const allExisting = [];
        Object.entries(availableTables.value).forEach(([cls, tbls]) => {
          tbls.forEach(t => allExisting.push({ ...t, cls }));
        });

        const comparisons = {};
        for (const tbl of (scanResult.value?.tables || [])) {
          let best = null, bestScore = 0;
          for (const existing of allExisting) {
            const res = await POST('/api/excel/compare-tables', {
              scanned_columns: tbl.columns.map(c => c.name),
              existing_table_id: existing.table_id,
            });
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
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
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
            tbls.forEach(t => suggestions.push({
              source_class: classId, source_table: t.table_name,
              target_table: t.table_name, where: '', cols: '', enabled: false,
            }));
          }
          collectMappings.value = suggestions;
        }
        step.value = 4;
      } else {
        step.value = 5;
      }
    };

    const register = async () => {
      if (!form.file_id) { emit('toast', { msg: 'ID du Post requis', type: 'error' }); return; }
      saving.value = true;
      try {
        const tblMappings = tableMappings.value.map(m => ({
          excel_name: m.scannedTable.name,
          schema_name: m.action === 'link' ? (m.linkedId?.split('.').pop() || m.schemaName) : m.schemaName,
          sheet: m.scannedTable.sheet,
          columns: m.columns.filter(c => c.include).map(c => ({
            name: c.name, col_type: c.col_type, header: c.name, is_key: c.is_key, required: c.required,
          })),
          include: true,
        }));
        const fieldMappings = (scanResult.value?.named_ranges || []).map(nr => ({
          named_range: nr.name, field_name: nr.name, field_type: nr.suggested_type,
          nature: 'identitaire', source: 'user_input', required: true, pushable: false, include: true,
        }));
        const activeCollects = collectMappings.value.filter(cm => cm.enabled).map(cm => ({
          source_class: cm.source_class, source_table: cm.source_table,
          target_table: cm.target_table, where_clause: cm.where, cols_filter: cm.cols,
        }));

        await POST('/api/excel/register', {
          path: uploadedFile.value?.name || '',
          file_id: form.file_id, type_fichier: selectedClass.value,
          owner_role: form.owner_role, synchro_periodicite: form.periodicite,
          field_mappings: fieldMappings, table_mappings: tblMappings,
          update_class_fields: form.update_class_fields, collect_mappings: activeCollects,
        });

        emit('toast', { msg: `Post '${form.file_id}' enregistré avec succès`, type: 'ok' });
        step.value = 1; uploadedFile.value = null; scanResult.value = null;
        selectedClass.value = ''; tableMappings.value = []; collectMappings.value = [];
        tableComparisons.value = {}; availableTables.value = {};
        Object.assign(form, { file_id: '', owner_role: '', periodicite: 'quotidien', update_class_fields: false });
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
      finally { saving.value = false; }
    };

    const verdictClass = v => v === 'identical' ? 'badge-green' : v === 'compatible' ? 'badge-orange' : 'badge-gray';
    const verdictLabel = s => s >= 0.9 ? `identique (${Math.round(s*100)}%)` : s >= 0.6 ? `compatible (${Math.round(s*100)}%)` : 'différent';

    return {
      step, scanning, saving, comparing, dragOver, uploadedFile, scanResult,
      fileTypes, actors, availableTables, selectedClass, tableComparisons,
      tableMappings, collectMappings, isCockpit, form,
      onDrop, onFileInput, handleFile, scanFile, compareAndProceed, proceedFromMapping,
      register, verdictClass, verdictLabel,
    };
  },
  template: `
    <div>
      <!-- Indicateurs d'étapes -->
      <div style="display:flex;align-items:center;gap:0;margin-bottom:24px;flex-wrap:wrap;gap:4px">
        <template v-for="(s, i) in ['Upload','Résultats','Mapping','COLLECT','Enregistrer']" :key="i">
          <div style="display:flex;align-items:center;gap:6px;padding:5px 12px;border-radius:6px;font-size:0.82rem;font-weight:500"
               :style="{background:step===i+1?'var(--accent)':'transparent',color:step===i+1?'#fff':'var(--text-dim)',opacity:(!isCockpit&&i===3)?0.35:1}">
            <span style="width:18px;height:18px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:0.72rem;font-weight:700"
                  :style="{background:step>i+1?'#22c55e':step===i+1?'rgba(255,255,255,0.25)':'var(--border)'}">
              {{ step>i+1?'✓':i+1 }}
            </span>
            {{ s }}
          </div>
          <div v-if="i<4" style="width:16px;height:1px;background:var(--border)"></div>
        </template>
      </div>

      <!-- ÉTAPE 1: Upload -->
      <div v-if="step===1" class="card">
        <div class="card-header"><span class="card-title">Choisir un fichier Excel</span></div>
        <div style="padding:16px">
          <div @dragover.prevent="dragOver=true" @dragleave.prevent="dragOver=false" @drop.prevent="onDrop"
               style="border:2px dashed;border-radius:12px;padding:48px 24px;text-align:center;cursor:pointer;transition:all 0.2s"
               :style="{borderColor:dragOver?'var(--accent)':'var(--border)',background:dragOver?'rgba(59,130,246,0.08)':'var(--surface-alt)'}"
               @click="$refs.fileInput.click()">
            <div style="font-size:2.5rem;margin-bottom:12px">📊</div>
            <div v-if="!uploadedFile" style="color:var(--text-dim);font-size:0.95rem">
              Glissez un fichier <strong>.xlsx</strong> ici<br>
              <span style="font-size:0.82rem">ou cliquez pour parcourir</span>
            </div>
            <div v-else style="color:#22c55e;font-weight:600">
              ✓ {{ uploadedFile.name }}<br>
              <span style="font-size:0.82rem;color:var(--text-dim)">{{ (uploadedFile.size/1024).toFixed(0) }} Ko</span>
            </div>
            <input ref="fileInput" type="file" accept=".xlsx,.xlsm,.xls" style="display:none" @change="onFileInput" />
          </div>
          <div style="margin-top:16px;display:flex;justify-content:flex-end">
            <button class="btn btn-primary" :disabled="!uploadedFile||scanning" @click="scanFile">
              {{ scanning ? '⏳ Scan en cours…' : '→ Scanner le fichier' }}
            </button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 2: Résultats du scan -->
      <div v-if="step===2" class="card">
        <div class="card-header"><span class="card-title">Résultats du scan</span></div>
        <div style="padding:16px">
          <div style="display:flex;gap:8px;margin-bottom:20px;flex-wrap:wrap">
            <span class="badge badge-gray">{{ scanResult?.path }}</span>
            <span class="badge badge-blue">{{ (scanResult?.sheets||[]).length }} feuilles</span>
            <span class="badge badge-purple">{{ (scanResult?.named_ranges||[]).length }} plages nommées</span>
            <span class="badge badge-green">{{ (scanResult?.tables||[]).length }} tables Excel</span>
          </div>

          <div v-if="(scanResult?.named_ranges||[]).length" style="margin-bottom:20px">
            <div style="font-weight:600;font-size:0.85rem;margin-bottom:8px;color:var(--text-dim)">PLAGES NOMMÉES</div>
            <table>
              <thead><tr><th>Nom</th><th>Feuille</th><th>Aperçu</th><th>Type</th></tr></thead>
              <tbody>
                <tr v-for="nr in scanResult.named_ranges" :key="nr.name">
                  <td><code style="color:var(--accent)">{{ nr.name }}</code></td>
                  <td style="color:var(--text-dim);font-size:0.82rem">{{ nr.sheet }}</td>
                  <td style="font-size:0.82rem">{{ nr.value_preview || '—' }}</td>
                  <td><span class="badge badge-gray" style="font-size:0.75rem">{{ nr.suggested_type }}</span></td>
                </tr>
              </tbody>
            </table>
          </div>

          <div v-if="(scanResult?.tables||[]).length" style="margin-bottom:20px">
            <div style="font-weight:600;font-size:0.85rem;margin-bottom:8px;color:var(--text-dim)">TABLES EXCEL</div>
            <table>
              <thead><tr><th>Nom</th><th>Feuille</th><th>Colonnes</th><th>Lignes</th></tr></thead>
              <tbody>
                <tr v-for="t in scanResult.tables" :key="t.name">
                  <td><code style="color:var(--accent)">{{ t.name }}</code></td>
                  <td style="color:var(--text-dim);font-size:0.82rem">{{ t.sheet }}</td>
                  <td>{{ t.columns.length }} col.</td>
                  <td style="color:var(--text-dim)">{{ t.row_count }}</td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="form-group" style="max-width:400px">
            <label>Classe (type de fichier) *</label>
            <select v-model="selectedClass">
              <option value="">— Choisir une Classe —</option>
              <option v-for="ft in fileTypes" :key="ft.id" :value="ft.id">{{ ft.id }} — {{ ft.label }}</option>
            </select>
          </div>

          <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px">
            <button class="btn btn-ghost" @click="step=1">← Retour</button>
            <button class="btn btn-primary" :disabled="!selectedClass||comparing" @click="compareAndProceed">
              {{ comparing ? '⏳ Comparaison…' : '→ Mapper les tables' }}
            </button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 3: Mapping des tables -->
      <div v-if="step===3" class="card">
        <div class="card-header"><span class="card-title">Mapping des tables</span></div>
        <div style="padding:16px">
          <div v-if="tableMappings.length===0" style="color:var(--text-dim);padding:24px;text-align:center">
            Aucune table trouvée dans ce fichier.
          </div>
          <div v-for="m in tableMappings" :key="m.scannedTable.name"
               style="border:1px solid var(--border);border-radius:8px;padding:14px;margin-bottom:12px">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap">
              <code style="color:var(--accent)">{{ m.scannedTable.name }}</code>
              <span v-if="m.comparison" class="badge" :class="verdictClass(m.comparison.verdict)">
                ⚡ {{ verdictLabel(m.comparison.match_score) }}
              </span>
              <span v-else class="badge badge-gray">Nouvelle</span>
              <select v-model="m.action" style="width:auto;padding:4px 8px;margin-left:auto">
                <option v-if="m.comparison" value="link">Lier à la table existante</option>
                <option value="create">Créer nouvelle table</option>
              </select>
            </div>
            <div v-if="m.action==='create'" style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
              <label style="font-size:0.82rem;white-space:nowrap">Nom schéma :</label>
              <input v-model="m.schemaName" style="width:200px" placeholder="TabNomTable" />
            </div>
            <div v-if="m.action==='link'&&m.comparison" style="font-size:0.82rem;color:var(--text-dim);margin-bottom:8px">
              Liée à : <code>{{ m.comparison.existing.table_id }}</code>
            </div>
            <div style="display:flex;flex-wrap:wrap;gap:5px">
              <span v-for="col in m.scannedTable.columns" :key="col.name"
                    style="font-size:0.77rem;padding:3px 7px;border-radius:4px;border:1px solid var(--border)"
                    :style="{background: m.comparison&&m.comparison.matching_columns.includes(col.name)?'rgba(34,197,94,0.1)':m.comparison&&m.comparison.extra_in_scan.includes(col.name)?'rgba(245,158,11,0.1)':'var(--surface-alt)'}">
                {{ col.name }} <span style="color:var(--text-dim)">· {{ col.suggested_type }}</span>
              </span>
            </div>
            <div v-if="m.comparison&&m.comparison.missing_in_scan.length" style="margin-top:6px;font-size:0.78rem;color:#f59e0b">
              ⚠ Colonnes manquantes : {{ m.comparison.missing_in_scan.join(', ') }}
            </div>
          </div>
          <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px">
            <button class="btn btn-ghost" @click="step=2">← Retour</button>
            <button class="btn btn-primary" @click="proceedFromMapping">
              {{ isCockpit ? '→ COLLECT Builder' : '→ Enregistrement' }}
            </button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 4: COLLECT Builder -->
      <div v-if="step===4" class="card">
        <div class="card-header"><span class="card-title">COLLECT Builder</span></div>
        <div style="padding:16px">
          <div style="margin-bottom:16px;padding:10px 14px;background:rgba(59,130,246,0.08);border:1px solid rgba(59,130,246,0.2);border-radius:6px;font-size:0.85rem;color:var(--text-dim)">
            Ce fichier (classe <strong>{{ selectedClass }}</strong>) peut agréger des tables depuis des fichiers enfants.
            Cochez les tables à collecter.
          </div>
          <div v-if="collectMappings.length===0" style="color:var(--text-dim);text-align:center;padding:24px">
            Aucune table disponible dans les classes enfants (tables.json vide pour ces classes).
          </div>
          <table v-else>
            <thead>
              <tr>
                <th style="width:28px"></th>
                <th>Source</th>
                <th>Dans ce fichier comme</th>
                <th>WHERE (optionnel)</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="cm in collectMappings" :key="cm.source_class+'_'+cm.source_table">
                <td><input type="checkbox" v-model="cm.enabled" style="width:auto" /></td>
                <td>
                  <span class="badge badge-gray" style="font-size:0.75rem">{{ cm.source_class }}</span>
                  <code style="color:var(--accent);margin-left:6px;font-size:0.82rem">{{ cm.source_table }}</code>
                </td>
                <td><input v-model="cm.target_table" :disabled="!cm.enabled" style="width:180px" /></td>
                <td><input v-model="cm.where" :disabled="!cm.enabled" placeholder="ex: statut = 'actif'" style="width:200px" /></td>
              </tr>
            </tbody>
          </table>
          <div style="display:flex;gap:8px;justify-content:space-between;margin-top:16px">
            <button class="btn btn-ghost" @click="step=3">← Retour</button>
            <button class="btn btn-primary" @click="step=5">→ Enregistrement</button>
          </div>
        </div>
      </div>

      <!-- ÉTAPE 5: Enregistrement -->
      <div v-if="step===5" class="card">
        <div class="card-header"><span class="card-title">Enregistrement</span></div>
        <div style="padding:16px">
          <div class="form-row">
            <div class="form-group">
              <label>ID du Post *</label>
              <input v-model="form.file_id" placeholder="ex: UO-006 ou L03U12-P042" />
            </div>
            <div class="form-group">
              <label>Classe</label>
              <input :value="selectedClass" disabled />
            </div>
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
                <option value="manuel">manuel</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>
              <input type="checkbox" v-model="form.update_class_fields" style="width:auto;margin-right:6px" />
              Mettre à jour les min_fields de la Classe avec les plages nommées
            </label>
          </div>

          <div style="background:var(--surface-alt);border:1px solid var(--border);border-radius:8px;padding:14px;margin:16px 0;font-size:0.85rem;line-height:1.8">
            <div style="font-weight:600;margin-bottom:6px;font-size:0.88rem">Résumé</div>
            <div style="color:#22c55e">✓ {{ (scanResult?.named_ranges||[]).length }} plages nommées → min_fields</div>
            <div style="color:#22c55e">✓ {{ tableMappings.filter(m=>m.action==='link').length }} table(s) liée(s)</div>
            <div style="color:#3b82f6">✓ {{ tableMappings.filter(m=>m.action==='create').length }} nouvelle(s) table(s) → tables.json</div>
            <div v-if="isCockpit" style="color:#8b5cf6">✓ {{ collectMappings.filter(cm=>cm.enabled).length }} COLLECT → hierarchy.json</div>
          </div>

          <div style="display:flex;gap:8px;justify-content:space-between">
            <button class="btn btn-ghost" @click="step=isCockpit?4:3">← Retour</button>
            <button class="btn btn-primary" :disabled="!form.file_id||saving" @click="register">
              {{ saving ? '⏳ Enregistrement…' : '✓ Enregistrer' }}
            </button>
          </div>
        </div>
      </div>
    </div>
  `
};

// ═══════════════════════════════════════════════════════════════════════════════
// ROOT APP
// ═══════════════════════════════════════════════════════════════════════════════
const App = {
  components: { ToastComponent, ViewDashboard, ViewFileTypes, ViewRegistry, ViewActors, ViewEcosystem, ViewHierarchy, ViewTables, ViewMxlPreview, ViewExcelImport, ViewSchemaBlueprint, ViewSchemaClasses, ViewSchemaRelations, ViewSchemaFunctions, ViewSchemaTemplates },
  setup() {
    const view = ref('dashboard');
    const selectedClassId = ref(null);
    const toasts = ref([]);
    const ecosystemStatus = ref(null);

    const addToast = ({ msg, type }) => {
      const id = Date.now();
      toasts.value.push({ id, msg, type });
      setTimeout(() => { toasts.value = toasts.value.filter(t => t.id !== id); }, 3500);
    };

    onMounted(async () => {
      ecosystemStatus.value = await GET('/api/ecosystem/status').catch(() => null);
    });

    const nav = [
      { id: 'dashboard',  icon: '⊞', label: 'Tableau de bord',    group: 'Principal' },
      { id: 'filetypes',  icon: '◈', label: 'Types de fichiers',   group: 'Configuration' },
      { id: 'registry',   icon: '◧', label: 'Registre',            group: 'Configuration' },
      { id: 'actors',     icon: '◉', label: 'Acteurs',             group: 'Configuration' },
      { id: 'excel-import', icon: '⤵', label: 'Importer Excel',     group: 'Configuration' },
      { id: 'hierarchy',  icon: '⬡', label: 'Hiérarchie',          group: 'Structure' },
      { id: 'tables',     icon: '▦', label: 'Tables & colonnes',   group: 'Structure' },
      { id: 'mxl',        icon: '⌥', label: 'Générateur MXL',      group: 'Structure' },
      { id: 'ecosystem',        icon: '◎', label: 'Graphe écosystème',   group: 'Visualisation' },
      { id: 'schema-blueprint', icon: '◈', label: 'Blueprint',           group: 'Schéma' },
      { id: 'schema-classes',   icon: '▣', label: 'Classes',             group: 'Schéma' },
      { id: 'schema-relations', icon: '↔', label: 'Relations P/F',       group: 'Schéma' },
      { id: 'schema-functions',  icon: '◎', label: 'Fonctions',           group: 'Schéma' },
      { id: 'schema-templates', icon: '⊟', label: 'Templates',           group: 'Schéma' },
    ];

    const navGroups = computed(() => {
      const g = {};
      nav.forEach(n => { if (!g[n.group]) g[n.group] = []; g[n.group].push(n); });
      return g;
    });

    const titles = {
      dashboard: 'Tableau de bord',
      filetypes: 'Types de fichiers',
      registry:  'Registre des fichiers',
      actors:    'Acteurs',
      hierarchy: 'Hiérarchie — LIST & COLLECT',
      tables:    'Tables & colonnes',
      mxl:       'Générateur de Manifeste MXL',
      'excel-import': 'Importer un fichier Excel',
      ecosystem:        'Graphe écosystème',
      'schema-blueprint': 'Schéma — Blueprint',
      'schema-classes':   'Schéma — Classes',
      'schema-relations': 'Schéma — Relations P/F',
      'schema-functions':  'Schéma — Fonctions',
      'schema-templates':  'Schéma — Templates',
    };

    const openClass = (classId) => {
      selectedClassId.value = classId;
      view.value = 'schema-classes';
    };

    return { view, selectedClassId, toasts, addToast, navGroups, titles, ecosystemStatus, openClass };
  },
  template: `
    <div id="sidebar">
      <div class="sidebar-logo">
        ExoSync<br><span>Studio · Ecosystem Designer</span>
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
        <span class="status-dot" :class="ecosystemStatus?.errors === 0 ? 'ok' : 'err'"></span>
        {{ ecosystemStatus ? ecosystemStatus.total_files + ' fichiers · ' + ecosystemStatus.edge_count + ' arêtes' : 'Chargement…' }}
      </div>
    </div>

    <div id="main">
      <div class="topbar">
        <span class="topbar-title">{{ titles[view] }}</span>
      </div>
      <div class="content">
        <view-dashboard   v-if="view==='dashboard'"  @toast="addToast" />
        <view-file-types  v-else-if="view==='filetypes'"  @toast="addToast" />
        <view-registry    v-else-if="view==='registry'"   @toast="addToast" />
        <view-actors      v-else-if="view==='actors'"     @toast="addToast" />
        <view-hierarchy   v-else-if="view==='hierarchy'"  @toast="addToast" />
        <view-tables      v-else-if="view==='tables'"     @toast="addToast" />
        <view-mxl-preview    v-else-if="view==='mxl'"           @toast="addToast" />
        <view-excel-import   v-else-if="view==='excel-import'" @toast="addToast" />
        <view-ecosystem          v-else-if="view==='ecosystem'"        @toast="addToast" />
        <view-schema-blueprint   v-else-if="view==='schema-blueprint'" @toast="addToast" @open-class="openClass" />
        <view-schema-classes     v-else-if="view==='schema-classes'"   @toast="addToast" :initial-class="selectedClassId" />
        <view-schema-relations   v-else-if="view==='schema-relations'" @toast="addToast" />
        <view-schema-functions   v-else-if="view==='schema-functions'"  @toast="addToast" />
        <view-schema-templates   v-else-if="view==='schema-templates'" @toast="addToast" />
      </div>
    </div>

    <toast-component :toasts="toasts" />
  `
};

createApp(App).mount('#app');
