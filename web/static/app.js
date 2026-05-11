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
// ROOT APP
// ═══════════════════════════════════════════════════════════════════════════════
const App = {
  components: { ToastComponent, ViewDashboard, ViewFileTypes, ViewRegistry, ViewActors, ViewEcosystem, ViewHierarchy, ViewTables, ViewMxlPreview, ViewSchemaBlueprint, ViewSchemaClasses, ViewSchemaRelations, ViewSchemaFunctions, ViewSchemaTemplates },
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
        <view-mxl-preview v-else-if="view==='mxl'"        @toast="addToast" />
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
