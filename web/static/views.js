// ── TagsInput (shared component) ─────────────────────────────────────────────
const TagsInput = {
  props: ['modelValue', 'placeholder'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const input = ref('');
    const tags = computed(() => props.modelValue || []);
    const add = () => {
      const v = input.value.trim();
      if (v && !tags.value.includes(v)) emit('update:modelValue', [...tags.value, v]);
      input.value = '';
    };
    const remove = t => emit('update:modelValue', tags.value.filter(x => x !== t));
    const onKey = e => { if (e.key === 'Enter' || e.key === ',') { e.preventDefault(); add(); } };
    return { input, tags, add, remove, onKey };
  },
  template: `
    <div class="tags-input" @click="$refs.inp.focus()">
      <span v-for="t in tags" :key="t" class="tag">
        {{ t }} <span class="tag-remove" @click.stop="remove(t)">×</span>
      </span>
      <input ref="inp" v-model="input" :placeholder="tags.length ? '' : (placeholder||'Entrée + Enter')"
             @keydown="onKey" @blur="add" />
    </div>
  `
};

// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Hierarchy (LIST + COLLECT)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewHierarchy = {
  emits: ['toast'],
  setup(_, { emit }) {
    const lists = ref([]);
    const collects = ref([]);
    const registry = ref([]);
    const activeTab = ref('lists');
    const showModal = ref(false);
    const modalType = ref('list');   // 'list' | 'collect'
    const editing = ref(null);

    const blankList = () => ({
      id: '', owner_file_id: '', list_name: '', form: 'TABLE',
      source_table: '', filter_type: '', filter_where: ''
    });
    const blankCollect = () => ({
      id: '', owner_file_id: '', source_table: '', list_name: '',
      target_table: '', where_clause: '', cols_filter: '', with_fields: ''
    });

    const formL = reactive(blankList());
    const formC = reactive(blankCollect());

    const load = async () => {
      [lists.value, collects.value, registry.value] = await Promise.all([
        GET('/api/hierarchy/lists'),
        GET('/api/hierarchy/collects'),
        GET('/api/registry'),
      ]);
    };
    onMounted(load);

    const genId = () => Math.random().toString(36).slice(2, 10);

    const openCreateList = () => {
      Object.assign(formL, blankList()); formL.id = genId();
      editing.value = null; modalType.value = 'list'; showModal.value = true;
    };
    const openEditList = item => {
      Object.assign(formL, { ...item });
      editing.value = item.id; modalType.value = 'list'; showModal.value = true;
    };
    const openCreateCollect = () => {
      Object.assign(formC, blankCollect()); formC.id = genId();
      editing.value = null; modalType.value = 'collect'; showModal.value = true;
    };
    const openEditCollect = item => {
      Object.assign(formC, { ...item });
      editing.value = item.id; modalType.value = 'collect'; showModal.value = true;
    };

    const saveList = async () => {
      try {
        if (editing.value) {
          await PUT(`/api/hierarchy/lists/${editing.value}`, formL);
          emit('toast', { msg: 'LIST mise à jour', type: 'ok' });
        } else {
          await POST('/api/hierarchy/lists', formL);
          emit('toast', { msg: 'LIST créée', type: 'ok' });
        }
        showModal.value = false; await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const saveCollect = async () => {
      try {
        if (editing.value) {
          await PUT(`/api/hierarchy/collects/${editing.value}`, formC);
          emit('toast', { msg: 'COLLECT mis à jour', type: 'ok' });
        } else {
          await POST('/api/hierarchy/collects', formC);
          emit('toast', { msg: 'COLLECT créé', type: 'ok' });
        }
        showModal.value = false; await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const delList = async id => {
      if (!confirm('Supprimer cette LIST ?')) return;
      try { await DEL(`/api/hierarchy/lists/${id}`); await load(); emit('toast', { msg: 'LIST supprimée', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };
    const delCollect = async id => {
      if (!confirm('Supprimer ce COLLECT ?')) return;
      try { await DEL(`/api/hierarchy/collects/${id}`); await load(); emit('toast', { msg: 'COLLECT supprimé', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const fileLabel = id => registry.value.find(f => f.id === id)?.id || id;

    return {
      lists, collects, registry, activeTab, showModal, modalType, editing,
      formL, formC, openCreateList, openEditList, openCreateCollect, openEditCollect,
      saveList, saveCollect, delList, delCollect, fileLabel
    };
  },
  template: `
    <div>
      <!-- Tabs -->
      <div style="display:flex;gap:0;margin-bottom:16px;border-bottom:1px solid var(--border)">
        <button class="btn btn-ghost" :style="activeTab==='lists'?'border-bottom:2px solid var(--accent);border-radius:0;color:var(--text)':'border-radius:0'"
                @click="activeTab='lists'">LIST <span class="topbar-badge">{{ lists.length }}</span></button>
        <button class="btn btn-ghost" :style="activeTab==='collects'?'border-bottom:2px solid var(--accent);border-radius:0;color:var(--text)':'border-radius:0'"
                @click="activeTab='collects'">COLLECT <span class="topbar-badge">{{ collects.length }}</span></button>
      </div>

      <!-- Lists tab -->
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
                <td><code style="color:var(--accent);font-size:0.78rem">{{ l.id }}</code></td>
                <td><strong>{{ l.owner_file_id }}</strong></td>
                <td><span class="badge badge-purple">{{ l.list_name }}</span></td>
                <td><span class="badge" :class="l.form==='TABLE'?'badge-blue':'badge-orange'">{{ l.form }}</span></td>
                <td style="font-size:0.8rem;color:var(--text-dim)">{{ l.source_table || '—' }}</td>
                <td style="font-size:0.78rem;color:var(--text-dim)">{{ l.filter_type || '—' }} {{ l.filter_where ? '= '+l.filter_where : '' }}</td>
                <td style="display:flex;gap:6px;justify-content:flex-end">
                  <button class="btn btn-ghost btn-sm" @click="openEditList(l)">Éditer</button>
                  <button class="btn btn-danger btn-sm" @click="delList(l.id)">✕</button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Collects tab -->
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
                <td><code style="color:var(--accent);font-size:0.78rem">{{ c.id }}</code></td>
                <td><strong>{{ c.owner_file_id }}</strong></td>
                <td><span class="badge badge-blue">{{ c.source_table }}</span></td>
                <td><span class="badge badge-purple">{{ c.list_name }}</span></td>
                <td><span class="badge badge-green">{{ c.target_table }}</span></td>
                <td style="font-size:0.78rem;color:var(--text-dim)">{{ c.where_clause || '—' }}</td>
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
            <div class="form-group">
              <label>ID</label>
              <input v-model="formL.id" :disabled="!!editing" />
            </div>
            <div class="form-group">
              <label>Forme</label>
              <select v-model="formL.form">
                <option value="TABLE">TABLE</option>
                <option value="DYNAMIC">DYNAMIC</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>Fichier père (owner) *</label>
            <select v-model="formL.owner_file_id">
              <option value="">— Choisir —</option>
              <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }} ({{ f.type_fichier }})</option>
            </select>
          </div>
          <div class="form-group">
            <label>Nom de la liste *</label>
            <input v-model="formL.list_name" placeholder="UOs_actifs" />
          </div>
          <div class="form-group">
            <label>Source table <span style="color:var(--text-dim)">(DYNAMIC)</span></label>
            <input v-model="formL.source_table" placeholder="TabUOs" />
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Type de filtre</label>
              <input v-model="formL.filter_type" placeholder="statut" />
            </div>
            <div class="form-group">
              <label>Condition WHERE</label>
              <input v-model="formL.filter_where" placeholder="actif" />
            </div>
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
            <label>Fichier père (owner) *</label>
            <select v-model="formC.owner_file_id">
              <option value="">— Choisir —</option>
              <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }} ({{ f.type_fichier }})</option>
            </select>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Table source *</label>
              <input v-model="formC.source_table" placeholder="Planning" />
            </div>
            <div class="form-group">
              <label>From list *</label>
              <input v-model="formC.list_name" placeholder="UOs_actifs" />
            </div>
          </div>
          <div class="form-group">
            <label>Into table (cible) *</label>
            <input v-model="formC.target_table" placeholder="vue_planning" />
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Clause WHERE</label>
              <input v-model="formC.where_clause" placeholder="statut != fermé" />
            </div>
            <div class="form-group">
              <label>Filtres colonnes</label>
              <input v-model="formC.cols_filter" placeholder="nom,date,avancement" />
            </div>
          </div>
          <div class="form-group">
            <label>WITH fields</label>
            <input v-model="formC.with_fields" placeholder="uo_id,uo_nom" />
          </div>
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
// VIEW: Tables + éditeur de colonnes
// ═══════════════════════════════════════════════════════════════════════════════
const COL_TYPES = ['string','float','int','date','pct','bool','KEY'];
const WRITE_MODES = ['','creation','engineer','admin'];

const ViewTables = {
  emits: ['toast'],
  setup(_, { emit }) {
    const tables = ref([]);
    const registry = ref([]);
    const selectedFile = ref('');
    const showModal = ref(false);
    const editing = ref(null);
    const importing = ref(false);

    const blankTable = () => ({
      id: '', file_id: '', table_name: '', sheet: '', description: '', columns: []
    });
    const form = reactive(blankTable());

    const load = async () => {
      registry.value = await GET('/api/registry');
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
      form.id = (selectedFile.value || 'file') + '.' + Date.now().toString(36);
      editing.value = null; showModal.value = true;
    };
    const openEdit = item => {
      Object.assign(form, JSON.parse(JSON.stringify(item)));
      editing.value = item.id; showModal.value = true;
    };

    const addColumn = () => {
      form.columns.push({ name: '', col_type: 'string', header: '', write: '', is_key: false, description: '' });
    };
    const removeColumn = i => form.columns.splice(i, 1);
    const moveUp = i => { if (i > 0) { const t = form.columns[i]; form.columns[i] = form.columns[i-1]; form.columns[i-1] = t; } };
    const moveDown = i => { if (i < form.columns.length-1) { const t = form.columns[i]; form.columns[i] = form.columns[i+1]; form.columns[i+1] = t; } };

    const save = async () => {
      try {
        if (editing.value) {
          await PUT(`/api/tables/${editing.value}`, form);
          emit('toast', { msg: 'Table mise à jour', type: 'ok' });
        } else {
          await POST('/api/tables', form);
          emit('toast', { msg: 'Table créée', type: 'ok' });
        }
        showModal.value = false; await loadTables();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const del = async id => {
      if (!confirm('Supprimer cette table ?')) return;
      try { await DEL(`/api/tables/${id}`); await loadTables(); emit('toast', { msg: 'Table supprimée', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const importFromEco = async () => {
      if (!selectedFile.value) return;
      importing.value = true;
      try {
        const imported = await GET(`/api/tables/from-ecosystem/${selectedFile.value}`);
        let count = 0;
        for (const t of imported) {
          try { await POST('/api/tables', t); count++; }
          catch(e) { /* already exists */ }
        }
        await loadTables();
        emit('toast', { msg: `${count} table(s) importée(s) depuis l'écosystème`, type: 'ok' });
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
      finally { importing.value = false; }
    };

    return {
      tables, registry, selectedFile, showModal, editing, importing,
      form, openCreate, openEdit, addColumn, removeColumn, moveUp, moveDown,
      save, del, importFromEco, COL_TYPES, WRITE_MODES
    };
  },
  template: `
    <div>
      <!-- Sélecteur de fichier -->
      <div class="card" style="margin-bottom:16px">
        <div style="display:flex;align-items:center;gap:12px">
          <label style="color:var(--text-dim);font-size:0.85rem;white-space:nowrap">Fichier :</label>
          <select v-model="selectedFile" style="flex:1;max-width:400px">
            <option value="">— Sélectionner un fichier —</option>
            <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }} — {{ f.type_fichier }}</option>
          </select>
          <button v-if="selectedFile" class="btn btn-ghost btn-sm" @click="importFromEco" :disabled="importing">
            {{ importing ? 'Import…' : '↓ Importer depuis l\'écosystème' }}
          </button>
          <button v-if="selectedFile" class="btn btn-primary" @click="openCreate">+ Nouvelle table</button>
        </div>
      </div>

      <div v-if="!selectedFile" class="empty">Sélectionner un fichier pour voir ses tables</div>

      <!-- Liste des tables -->
      <div v-for="tbl in tables" :key="tbl.id" class="card">
        <div class="card-header">
          <div>
            <span class="card-title">{{ tbl.table_name }}</span>
            <span style="color:var(--text-dim);font-size:0.8rem;margin-left:8px">SHEET={{ tbl.sheet }}</span>
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
              <td><span class="badge" :class="col.col_type==='KEY'?'badge-orange':col.col_type==='float'||col.col_type==='int'?'badge-blue':'badge-gray'">{{ col.col_type }}</span></td>
              <td style="color:var(--text-dim)">{{ col.header || '—' }}</td>
              <td style="font-size:0.78rem;color:var(--text-dim)">{{ col.write || '—' }}</td>
              <td>{{ col.is_key ? '✓' : '' }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Modal éditeur de table -->
      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal" style="width:700px;max-width:98vw">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvelle' }} table</div>
          <div class="form-row">
            <div class="form-group">
              <label>Nom de la table *</label>
              <input v-model="form.table_name" placeholder="TabActivites" />
            </div>
            <div class="form-group">
              <label>Feuille Excel *</label>
              <input v-model="form.sheet" placeholder="Activites" />
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <input v-model="form.description" placeholder="…" />
          </div>

          <!-- Colonnes -->
          <div style="margin-top:16px;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between">
            <span style="font-size:0.85rem;font-weight:600">Colonnes</span>
            <button class="btn btn-ghost btn-sm" @click="addColumn">+ Colonne</button>
          </div>
          <div style="max-height:340px;overflow-y:auto">
            <table>
              <thead>
                <tr>
                  <th style="width:28px"></th>
                  <th>Nom</th><th>Type</th><th>Header</th><th>Write</th><th>Clé</th>
                  <th style="width:36px"></th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(col,i) in form.columns" :key="i">
                  <td>
                    <div style="display:flex;flex-direction:column;gap:1px">
                      <span style="cursor:pointer;font-size:0.7rem;color:var(--text-dim)" @click="moveUp(i)">▲</span>
                      <span style="cursor:pointer;font-size:0.7rem;color:var(--text-dim)" @click="moveDown(i)">▼</span>
                    </div>
                  </td>
                  <td><input v-model="col.name" placeholder="nom" style="min-width:80px" /></td>
                  <td>
                    <select v-model="col.col_type" style="min-width:70px">
                      <option v-for="t in COL_TYPES" :key="t" :value="t">{{ t }}</option>
                    </select>
                  </td>
                  <td><input v-model="col.header" placeholder="Libellé" /></td>
                  <td>
                    <select v-model="col.write" style="min-width:80px">
                      <option v-for="w in WRITE_MODES" :key="w" :value="w">{{ w || '(none)' }}</option>
                    </select>
                  </td>
                  <td style="text-align:center">
                    <input type="checkbox" v-model="col.is_key" style="width:auto" />
                  </td>
                  <td>
                    <button class="btn btn-danger btn-sm" @click="removeColumn(i)">✕</button>
                  </td>
                </tr>
                <tr v-if="form.columns.length===0">
                  <td colspan="7" style="text-align:center;color:var(--text-dim);padding:16px">Aucune colonne — cliquez "+ Colonne"</td>
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
// VIEW: MXL Preview
// ═══════════════════════════════════════════════════════════════════════════════
const ViewMxlPreview = {
  emits: ['toast'],
  setup(_, { emit }) {
    const registry = ref([]);
    const selectedFile = ref('');
    const mxl = ref('');
    const loading = ref(false);

    onMounted(async () => { registry.value = await GET('/api/registry'); });

    const generate = async () => {
      if (!selectedFile.value) return;
      loading.value = true; mxl.value = '';
      try {
        const res = await GET(`/api/mxl/${selectedFile.value}`);
        mxl.value = res.mxl;
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
      finally { loading.value = false; }
    };
    watch(selectedFile, generate);

    const copy = async () => {
      await navigator.clipboard.writeText(mxl.value);
      emit('toast', { msg: 'MXL copié dans le presse-papiers', type: 'ok' });
    };

    const download = () => {
      const blob = new Blob([mxl.value], { type: 'text/plain' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `manifeste_${selectedFile.value}.mxl`;
      a.click();
    };

    // Syntax highlighting simple
    const highlighted = computed(() => {
      if (!mxl.value) return '';
      return mxl.value
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/^(FILE_TYPE|FILE_ID|VERSION)(\s+.*)$/gm,
          '<span style="color:#f97316;font-weight:600">$1</span><span style="color:#e2e8f0">$2</span>')
        .replace(/^(DEF|COL|PULL|PUSH|LIST|COLLECT|COMPUTE|BIND)(\s+)/gm,
          '<span style="color:#60a5fa;font-weight:600">$1</span><span style="color:#e2e8f0">$2</span>')
        .replace(/\b(TYPE|SHEET|HEADER|WRITE|KEY|FROM|TO|FORM|SOURCE|FILTER|WHERE|MODE|FORMULA|FROM_LIST|INTO|COLS|WITH)=/g,
          '<span style="color:#a78bfa">$1</span>=')
        .replace(/"([^"]*)"/g, '<span style="color:#34d399">"$1"</span>')
        .replace(/^(#.*)$/gm, '<span style="color:#475569;font-style:italic">$1</span>');
    });

    return { registry, selectedFile, mxl, loading, highlighted, generate, copy, download };
  },
  template: `
    <div>
      <div class="card" style="margin-bottom:16px">
        <div style="display:flex;align-items:center;gap:12px">
          <label style="color:var(--text-dim);font-size:0.85rem;white-space:nowrap">Fichier :</label>
          <select v-model="selectedFile" style="flex:1;max-width:400px">
            <option value="">— Sélectionner un fichier —</option>
            <option v-for="f in registry" :key="f.id" :value="f.id">{{ f.id }} — {{ f.type_fichier }}</option>
          </select>
          <button class="btn btn-ghost btn-sm" @click="generate" :disabled="!selectedFile||loading">
            {{ loading ? '⌛ Génération…' : '↺ Regénérer' }}
          </button>
          <button v-if="mxl" class="btn btn-ghost btn-sm" @click="copy">⎘ Copier</button>
          <button v-if="mxl" class="btn btn-primary btn-sm" @click="download">↓ .mxl</button>
        </div>
      </div>

      <div v-if="!selectedFile" class="empty">Sélectionner un fichier pour générer son Manifeste MXL</div>

      <div v-if="mxl" class="card" style="padding:0">
        <div style="padding:12px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between">
          <span style="font-size:0.8rem;color:var(--text-dim)">manifeste_{{ selectedFile }}.mxl</span>
          <span class="topbar-badge">{{ mxl.split('\\n').length }} lignes</span>
        </div>
        <pre style="padding:16px;font-size:0.82rem;line-height:1.6;overflow-x:auto;tab-size:2;font-family:'Cascadia Code','Fira Code',monospace"
             v-html="highlighted"></pre>
      </div>
    </div>
  `
};
