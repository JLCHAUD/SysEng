// ═══════════════════════════════════════════════════════════════════════════════
// SCHEMA VIEWS — Niveau 1 : Schéma d'écosystème
// ═══════════════════════════════════════════════════════════════════════════════

// Couleurs par fonction owner — alignées sur SCHEMA-01 (interne / externe)
const FUNC_COLORS = {
  // Fonctions internes
  pilote_pole:          { bg: '#431407', border: '#f97316', text: '#fb923c' },
  pilote_metier:        { bg: '#14532d', border: '#22c55e', text: '#4ade80' },
  ingenieur_sys:        { bg: '#1e3a5f', border: '#3b82f6', text: '#60a5fa' },
  it_manager:           { bg: '#3b0764', border: '#a855f7', text: '#c084fc' },
  // Fonctions externes — donneurs d'ordre
  donneur_ordre_pole:   { bg: '#292524', border: '#d97706', text: '#fbbf24' },
  donneur_ordre_metier: { bg: '#1c1917', border: '#78716c', text: '#a8a29e' },
  // Fonctions externes — correspondants techniques
  ingenieur_sys_client: { bg: '#134e4a', border: '#14b8a6', text: '#2dd4bf' },
  expert_technique:     { bg: '#4a044e', border: '#ec4899', text: '#f472b6' },
  fournisseur:          { bg: '#1e293b', border: '#64748b', text: '#94a3b8' },
  correspondant_projet: { bg: '#1e1b4b', border: '#6366f1', text: '#818cf8' },
};
const DEFAULT_COLOR = { bg: '#1e293b', border: '#475569', text: '#94a3b8' };
const funcColor = id => FUNC_COLORS[id] || DEFAULT_COLOR;

// ── Composant : éditeur de colonnes inline ─────────────────────────────────────
const ColEditor = {
  props: ['modelValue'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const cols = computed(() => props.modelValue || []);
    const add = () => emit('update:modelValue', [...cols.value,
      { name:'', col_type:'string', header:'', write:'', is_key:false, required:true, description:'' }]);
    const remove = i => emit('update:modelValue', cols.value.filter((_,j) => j!==i));
    const update = (i, field, val) => {
      const next = cols.value.map((c,j) => j===i ? {...c, [field]:val} : c);
      emit('update:modelValue', next);
    };
    return { cols, add, remove, update, COL_TYPES, WRITE_MODES };
  },
  template: `
    <div>
      <div style="display:flex;justify-content:flex-end;margin-bottom:6px">
        <button class="btn btn-ghost btn-sm" @click="add">+ Colonne</button>
      </div>
      <table v-if="cols.length">
        <thead>
          <tr>
            <th>Nom</th><th>Type</th><th>Header</th><th>Write</th>
            <th title="Clé primaire">Clé</th><th title="Minimum requis">Min</th><th></th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(col,i) in cols" :key="i">
            <td><input :value="col.name" @input="update(i,'name',$event.target.value)" placeholder="nom_col" /></td>
            <td>
              <select :value="col.col_type" @change="update(i,'col_type',$event.target.value)" style="min-width:70px">
                <option v-for="t in COL_TYPES" :key="t" :value="t">{{ t }}</option>
              </select>
            </td>
            <td><input :value="col.header" @input="update(i,'header',$event.target.value)" placeholder="Libellé" /></td>
            <td>
              <select :value="col.write" @change="update(i,'write',$event.target.value)" style="min-width:80px">
                <option v-for="w in WRITE_MODES" :key="w" :value="w">{{ w||'—' }}</option>
              </select>
            </td>
            <td style="text-align:center">
              <input type="checkbox" :checked="col.is_key" @change="update(i,'is_key',$event.target.checked)" style="width:auto" />
            </td>
            <td style="text-align:center">
              <input type="checkbox" :checked="col.required" @change="update(i,'required',$event.target.checked)" style="width:auto" />
            </td>
            <td><button class="btn btn-danger btn-sm" @click="remove(i)">✕</button></td>
          </tr>
        </tbody>
      </table>
      <div v-else style="color:var(--text-dim);font-size:0.8rem;padding:8px">
        Aucune colonne — cliquez "+ Colonne"
      </div>
    </div>
  `
};

// ── Composant : éditeur de tables standard ─────────────────────────────────────
const TableStdEditor = {
  components: { ColEditor },
  props: ['modelValue'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const tables = computed(() => props.modelValue || []);
    const expanded = ref(null);
    const add = () => {
      const next = [...tables.value, { name:'', sheet:'', columns:[], description:'' }];
      emit('update:modelValue', next);
      expanded.value = next.length - 1;
    };
    const remove = i => {
      if (expanded.value === i) expanded.value = null;
      emit('update:modelValue', tables.value.filter((_,j) => j!==i));
    };
    const update = (i, field, val) => {
      emit('update:modelValue', tables.value.map((t,j) => j===i ? {...t,[field]:val} : t));
    };
    const updateCols = (i, cols) => update(i, 'columns', cols);
    const toggle = i => { expanded.value = expanded.value === i ? null : i; };
    return { tables, expanded, add, remove, update, updateCols, toggle };
  },
  template: `
    <div>
      <div v-for="(tbl,i) in tables" :key="i" style="border:1px solid var(--border);border-radius:6px;margin-bottom:8px">
        <div style="display:flex;align-items:center;gap:8px;padding:10px 12px;cursor:pointer" @click="toggle(i)">
          <span style="flex:1;font-size:0.85rem;font-weight:600">
            {{ tbl.name || '(sans nom)' }}
            <span style="color:var(--text-dim);font-weight:400;font-size:0.75rem"> SHEET={{ tbl.sheet||'?' }}</span>
            <span class="topbar-badge" style="margin-left:6px">{{ tbl.columns.length }} col.</span>
          </span>
          <span style="color:var(--text-dim);font-size:0.8rem">{{ expanded===i ? '▲' : '▼' }}</span>
          <button class="btn btn-danger btn-sm" @click.stop="remove(i)">✕</button>
        </div>
        <div v-if="expanded===i" style="padding:0 12px 12px;border-top:1px solid var(--border)">
          <div class="form-row" style="margin-top:10px">
            <div class="form-group">
              <label>Nom de la table *</label>
              <input :value="tbl.name" @input="update(i,'name',$event.target.value)" placeholder="TabActivites" />
            </div>
            <div class="form-group">
              <label>Feuille Excel *</label>
              <input :value="tbl.sheet" @input="update(i,'sheet',$event.target.value)" placeholder="Activites" />
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <input :value="tbl.description" @input="update(i,'description',$event.target.value)" placeholder="Rôle de cette table…" />
          </div>
          <div style="margin-top:8px">
            <label style="font-size:0.78rem;color:var(--text-dim);margin-bottom:6px;display:block">
              Colonnes — <em>Min = colonne obligatoire dans tout Post de cette Classe</em>
            </label>
            <col-editor :modelValue="tbl.columns" @update:modelValue="updateCols(i,$event)" />
          </div>
        </div>
      </div>
      <button class="btn btn-ghost btn-sm" @click="add" style="width:100%;margin-top:4px">+ Table standard</button>
    </div>
  `
};

// ── Constantes champs scalaires ───────────────────────────────────────────────
const FIELD_TYPES   = ['string','int','float','date','bool','ref'];
const FIELD_NATURES = ['identitaire','operationnel','parametre'];
const FIELD_SOURCES = ['user_input','system','computed','reference'];

// ── Composant : éditeur de champs scalaires (min_fields) ──────────────────────
const FieldEditor = {
  props: ['modelValue'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const fields = computed(() => props.modelValue || []);
    const add = () => emit('update:modelValue', [...fields.value,
      { name:'', label:'', field_type:'string', nature:'identitaire',
        source:'user_input', required:true, pushable:false, ref_class:'', description:'' }]);
    const remove = i => emit('update:modelValue', fields.value.filter((_,j) => j!==i));
    const update = (i, key, val) => {
      emit('update:modelValue', fields.value.map((f,j) => j===i ? {...f, [key]:val} : f));
    };
    const natureBadge = n => n==='operationnel' ? 'badge-orange' : n==='parametre' ? 'badge-purple' : 'badge-blue';
    return { fields, add, remove, update, FIELD_TYPES, FIELD_NATURES, FIELD_SOURCES, natureBadge };
  },
  template: `
    <div>
      <div style="display:flex;justify-content:flex-end;margin-bottom:6px">
        <button class="btn btn-ghost btn-sm" @click="add">+ Champ</button>
      </div>
      <div v-if="!fields.length" style="color:var(--text-dim);font-size:0.8rem;padding:8px">
        Aucun champ scalaire — cliquez "+ Champ"
      </div>
      <div v-for="(f,i) in fields" :key="i"
           style="border:1px solid var(--border);border-radius:6px;padding:10px;margin-bottom:8px;background:var(--surface2)">
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:6px">
          <div>
            <label style="font-size:0.7rem">Nom *</label>
            <input :value="f.name" @input="update(i,'name',$event.target.value)" placeholder="avancement_global" />
          </div>
          <div>
            <label style="font-size:0.7rem">Label</label>
            <input :value="f.label" @input="update(i,'label',$event.target.value)" placeholder="Avancement global" />
          </div>
          <div>
            <label style="font-size:0.7rem">Type</label>
            <select :value="f.field_type" @change="update(i,'field_type',$event.target.value)">
              <option v-for="t in FIELD_TYPES" :key="t" :value="t">{{ t }}</option>
            </select>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr auto auto auto;gap:8px;align-items:end">
          <div>
            <label style="font-size:0.7rem">Nature</label>
            <select :value="f.nature" @change="update(i,'nature',$event.target.value)">
              <option v-for="n in FIELD_NATURES" :key="n" :value="n">{{ n }}</option>
            </select>
          </div>
          <div>
            <label style="font-size:0.7rem">Source</label>
            <select :value="f.source" @change="update(i,'source',$event.target.value)">
              <option v-for="s in FIELD_SOURCES" :key="s" :value="s">{{ s }}</option>
            </select>
          </div>
          <div style="text-align:center">
            <label style="font-size:0.7rem;display:block">Requis</label>
            <input type="checkbox" :checked="f.required" @change="update(i,'required',$event.target.checked)" style="width:auto" />
          </div>
          <div style="text-align:center">
            <label style="font-size:0.7rem;display:block">Push</label>
            <input type="checkbox" :checked="f.pushable" @change="update(i,'pushable',$event.target.checked)" style="width:auto" />
          </div>
          <button class="btn btn-danger btn-sm" @click="remove(i)">✕</button>
        </div>
        <div v-if="f.source==='reference'" style="margin-top:6px">
          <label style="font-size:0.7rem">Classe référencée</label>
          <input :value="f.ref_class" @input="update(i,'ref_class',$event.target.value)" placeholder="uo_instance" />
        </div>
      </div>
    </div>
  `
};

// ═══════════════════════════════════════════════════════════════════════════════
// WIZARD de création de Classe (5 étapes)
// ═══════════════════════════════════════════════════════════════════════════════
const ClassWizard = {
  components: { TagsInput, FieldEditor, TableStdEditor },
  props: ['functions', 'editData'],
  emits: ['save', 'cancel'],
  setup(props, { emit }) {
    const step = ref(1);
    const STEPS = ['Identité', 'Structure', 'Champs', 'Tables', 'Aperçu'];
    const isEdit = computed(() => !!props.editData);

    const blank = () => ({
      id: '', label: '', description: '', owner_function: '',
      min_sheets: [], optional_sheets: [],
      allowed_namespaces: [], push_prefix: '',
      template: '', min_fields: [], std_tables: [],
    });

    const form = reactive(blank());
    watch(() => props.editData, d => {
      if (d) Object.assign(form, JSON.parse(JSON.stringify(d)));
      else Object.assign(form, blank());
      step.value = 1;
    }, { immediate: true });

    const mxlPreview = computed(() => {
      const lines = [
        `FILE_TYPE  ${form.id || '…'}`,
        `FILE_ID    {id}`,
        `VERSION    1`,
        '',
      ];
      for (const fld of form.min_fields) {
        if (!fld.name) continue;
        const parts = [`DEF  ${fld.name}`, `TYPE=${fld.field_type}`];
        if (fld.label) parts.push(`LABEL="${fld.label}"`);
        if (fld.pushable) parts.push('PUSH');
        lines.push(parts.join('  '));
      }
      if (form.min_fields.length) lines.push('');
      for (const tbl of form.std_tables) {
        if (!tbl.name) continue;
        lines.push(`DEF  ${tbl.name}  SHEET=${tbl.sheet}`);
        for (const col of tbl.columns) {
          const parts = [`  COL  ${col.name}`, `TYPE=${col.col_type}`];
          if (col.header) parts.push(`HEADER="${col.header}"`);
          if (col.write)  parts.push(`WRITE=${col.write}`);
          if (col.is_key) parts.push('KEY');
          lines.push(parts.join('  '));
        }
        lines.push('');
      }
      if (form.allowed_namespaces.length || form.push_prefix) {
        lines.push(`# Namespaces autorisés : ${form.allowed_namespaces.join(' ')}`);
        lines.push(`# Push prefix : ${form.push_prefix}`);
      }
      return lines.join('\n').trim();
    });

    const canNext = computed(() => {
      if (step.value === 1) return form.id.trim() && form.label.trim() && form.owner_function;
      return true;
    });

    const save = () => emit('save', { ...form });
    const cancel = () => emit('cancel');

    return { step, STEPS, form, isEdit, mxlPreview, canNext, save, cancel };
  },
  template: `
    <div class="modal-overlay" @click.self="cancel">
      <div class="modal" style="width:680px;max-width:98vw">

        <!-- Étapes -->
        <div style="display:flex;gap:0;margin-bottom:24px;border-bottom:1px solid var(--border);padding-bottom:16px">
          <div v-for="(s,i) in STEPS" :key="i"
               style="flex:1;text-align:center;font-size:0.78rem;cursor:pointer"
               :style="step===i+1 ? 'color:var(--accent);font-weight:600' : 'color:var(--text-dim)'"
               @click="step=i+1">
            <div style="width:24px;height:24px;border-radius:50%;margin:0 auto 4px;display:flex;align-items:center;justify-content:center;font-size:0.75rem"
                 :style="step===i+1 ? 'background:var(--accent);color:#fff' :
                         step>i+1  ? 'background:#22c55e;color:#fff' :
                                     'background:var(--surface2);color:var(--text-dim)'">
              {{ step > i+1 ? '✓' : i+1 }}
            </div>
            {{ s }}
          </div>
        </div>

        <!-- Étape 1 — Identité -->
        <div v-if="step===1">
          <div class="modal-title" style="margin-bottom:16px">Identité de la Classe</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID unique *</label>
              <input v-model="form.id" :disabled="isEdit" placeholder="uo_instance" />
              <div style="font-size:0.72rem;color:var(--text-dim);margin-top:3px">Identifiant technique — pas d'espaces</div>
            </div>
            <div class="form-group">
              <label>Fonction owner *</label>
              <select v-model="form.owner_function">
                <option value="">— Choisir —</option>
                <option v-for="f in functions" :key="f.id" :value="f.id">{{ f.label }}</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>Label *</label>
            <input v-model="form.label" placeholder="Unité d'Œuvre (instance)" />
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="form.description" rows="2" placeholder="Ce que représente cette Classe dans l'écosystème…" />
          </div>
        </div>

        <!-- Étape 2 — Structure -->
        <div v-if="step===2">
          <div class="modal-title" style="margin-bottom:4px">Structure minimale</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px">
            Les Posts de cette Classe peuvent avoir <em>plus</em> que ce minimum.
          </div>
          <div class="form-group">
            <label>Feuilles minimales <span style="color:var(--text-dim)">(toujours présentes)</span></label>
            <tags-input v-model="form.min_sheets" placeholder="_Manifeste, Activites…" />
          </div>
          <div class="form-group">
            <label>Feuilles optionnelles <span style="color:var(--text-dim)">(peuvent être présentes)</span></label>
            <tags-input v-model="form.optional_sheets" placeholder="Dashboard, _Log…" />
          </div>
          <div class="form-row" style="margin-top:4px">
            <div class="form-group">
              <label>Namespaces Store autorisés</label>
              <tags-input v-model="form.allowed_namespaces" placeholder="uo. ref. projet.…" />
            </div>
            <div class="form-group">
              <label>Push prefix</label>
              <input v-model="form.push_prefix" placeholder="uo.{id}." />
            </div>
          </div>
        </div>

        <!-- Étape 3 — Champs scalaires -->
        <div v-if="step===3">
          <div class="modal-title" style="margin-bottom:4px">Champs scalaires (DEF)</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px">
            Données unitaires qui caractérisent chaque Post — identité, état opérationnel ou paramètre.<br>
            Ces champs correspondent aux déclarations <code>DEF</code> du <code>_Manifeste</code>.
          </div>
          <div style="max-height:420px;overflow-y:auto">
            <field-editor v-model="form.min_fields" />
          </div>
        </div>

        <!-- Étape 4 — Tables standard -->
        <div v-if="step===4">
          <div class="modal-title" style="margin-bottom:4px">Tables standard</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px">
            Définissez le minimum de tables attendu pour tout Post de cette Classe.<br>
            Les Posts peuvent ajouter des tables supplémentaires.<br>
            <strong>Min</strong> sur les colonnes = colonne obligatoire dans tout Post.
          </div>
          <div style="max-height:380px;overflow-y:auto">
            <table-std-editor v-model="form.std_tables" />
          </div>
        </div>

        <!-- Étape 5 — Aperçu MXL -->
        <div v-if="step===5">
          <div class="modal-title" style="margin-bottom:4px">Aperçu du template MXL</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:12px">
            Ce template sera utilisé comme base pour les nouveaux Posts de cette Classe.
          </div>
          <pre style="background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:14px;font-size:0.8rem;line-height:1.6;max-height:320px;overflow-y:auto;font-family:monospace">{{ mxlPreview }}</pre>
          <div style="margin-top:12px;padding:10px 12px;background:var(--surface2);border-radius:6px;font-size:0.8rem;color:var(--text-dim)">
            ✓ <strong>{{ form.id }}</strong> — {{ form.min_fields.length }} champ(s) scalaire(s),
            {{ form.std_tables.length }} table(s) standard,
            {{ form.min_sheets.length }} feuille(s) minimale(s)
          </div>
        </div>

        <!-- Navigation -->
        <div class="form-actions" style="margin-top:20px">
          <button class="btn btn-ghost" @click="cancel">Annuler</button>
          <div style="flex:1"></div>
          <button v-if="step>1" class="btn btn-ghost" @click="step--">← Précédent</button>
          <button v-if="step<5" class="btn btn-primary" @click="step++" :disabled="!canNext">
            Suivant →
          </button>
          <button v-if="step===5" class="btn btn-primary" @click="save">
            {{ isEdit ? 'Enregistrer' : 'Créer la Classe' }}
          </button>
        </div>
      </div>
    </div>
  `
};

// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Schema Blueprint
// ═══════════════════════════════════════════════════════════════════════════════
const ViewSchemaBlueprint = {
  emits: ['toast', 'open-class'],
  setup(_, { emit }) {
    const blueprint = ref(null);
    const cyInstance = ref(null);
    const selected = ref(null);
    const showWizard = ref(false);
    const functions = ref([]);

    const load = async () => {
      [blueprint.value, functions.value] = await Promise.all([
        GET('/api/schema/blueprint'),
        GET('/api/schema/functions'),
      ]);
    };

    onMounted(async () => {
      await load();
      await nextTick();
      buildGraph();
    });

    const buildGraph = () => {
      if (!blueprint.value || typeof cytoscape === 'undefined') return;
      const { classes, relations } = blueprint.value;

      const elements = [];

      classes.forEach(cls => {
        const c = funcColor(cls.owner_function);
        elements.push({
          data: {
            id: cls.id, label: cls.label, fn: cls.owner_function,
            post_count: cls.post_count, tables: cls.std_tables?.length || 0,
            color: c.border, bg: c.bg,
          }
        });
      });

      relations.forEach(rel => {
        elements.push({
          data: {
            id: rel.id, source: rel.parent_class, target: rel.child_class,
            qualifier: rel.qualifier, cardinality: rel.cardinality,
            color: rel.qualifier === 'PRESCRIBED' ? '#ef4444' : '#94a3b8',
            linestyle: rel.qualifier === 'PRESCRIBED' ? 'solid' : 'dashed',
          }
        });
      });

      const cy = cytoscape({
        container: document.getElementById('schema-cy'),
        elements,
        style: [
          {
            selector: 'node',
            style: {
              'background-color': 'data(bg)',
              'border-color': 'data(color)',
              'border-width': 2,
              'label': 'data(label)',
              'color': '#e2e8f0',
              'font-size': '12px',
              'text-valign': 'center',
              'text-halign': 'center',
              'text-wrap': 'wrap',
              'text-max-width': '100px',
              'width': 120, 'height': 50,
              shape: 'roundrectangle',
            }
          },
          {
            selector: 'node:selected',
            style: { 'border-width': 3, 'border-color': '#fff' }
          },
          {
            selector: 'node.incomplete',
            style: { 'border-style': 'dashed' }
          },
          {
            selector: 'edge',
            style: {
              'line-color': 'data(color)',
              'target-arrow-color': 'data(color)',
              'target-arrow-shape': 'triangle',
              'curve-style': 'bezier',
              'width': 2,
              'line-style': 'data(linestyle)',
              'label': 'data(cardinality)',
              'font-size': '10px',
              'color': '#94a3b8',
              'text-background-color': '#0f172a',
              'text-background-opacity': 1,
              'text-background-padding': '2px',
            }
          },
          { selector: '.faded', style: { opacity: 0.2 } }
        ],
        layout: {
          name: 'breadthfirst',
          directed: true,
          padding: 40,
          spacingFactor: 1.4,
          roots: classes
            .filter(c => !relations.some(r => r.child_class === c.id))
            .map(c => c.id),
        }
      });

      // Marquer les classes incomplètes
      cy.nodes().forEach(n => {
        const cls = classes.find(c => c.id === n.id());
        if (cls && cls.std_tables.length === 0) n.addClass('incomplete');
      });

      cy.on('tap', 'node', evt => {
        const id = evt.target.id();
        selected.value = classes.find(c => c.id === id) || null;
        cy.elements().addClass('faded');
        evt.target.removeClass('faded');
        evt.target.neighborhood().removeClass('faded');
      });
      cy.on('tap', evt => {
        if (evt.target === cy) { selected.value = null; cy.elements().removeClass('faded'); }
      });
      cy.on('dblclick', 'node', evt => {
        emit('open-class', evt.target.id());
      });

      cyInstance.value = cy;
    };

    const saveClass = async (form) => {
      try {
        await POST('/api/schema/classes', form);
        emit('toast', { msg: 'Classe créée', type: 'ok' });
        showWizard.value = false;
        await load();
        await nextTick();
        buildGraph();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const fit = () => cyInstance.value?.fit(40);
    const relayout = () => {
      if (!cyInstance.value || !blueprint.value) return;
      cyInstance.value.layout({
        name: 'breadthfirst', directed: true, padding: 40, spacingFactor: 1.4,
        roots: blueprint.value.classes
          .filter(c => !blueprint.value.relations.some(r => r.child_class === c.id))
          .map(c => c.id),
      }).run();
    };

    return { blueprint, selected, showWizard, functions, saveClass, fit, relayout, funcColor };
  },
  template: `
    <div style="position:relative">
      <!-- Toolbar -->
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
        <div v-if="blueprint" style="font-size:0.82rem;color:var(--text-dim)">
          {{ blueprint.classes.length }} Classes · {{ blueprint.relations.length }} Relations
        </div>
        <div style="flex:1"></div>
        <button class="btn btn-ghost btn-sm" @click="fit">⊡ Recadrer</button>
        <button class="btn btn-ghost btn-sm" @click="relayout">↺ Relayout</button>
        <button class="btn btn-primary" @click="showWizard=true">+ Nouvelle Classe</button>
      </div>

      <!-- Légende -->
      <div style="display:flex;gap:16px;margin-bottom:10px;font-size:0.75rem;color:var(--text-dim);flex-wrap:wrap">
        <span>Double-clic → fiche détaillée</span>
        <span style="display:flex;align-items:center;gap:4px">
          <span style="width:20px;height:2px;background:#ef4444;display:inline-block"></span> PRESCRIBED
        </span>
        <span style="display:flex;align-items:center;gap:4px">
          <span style="width:20px;height:2px;background:#94a3b8;display:inline-block;border-top:2px dashed #94a3b8"></span> TYPICAL
        </span>
        <span style="display:flex;align-items:center;gap:4px">
          <span style="width:14px;height:14px;border:2px dashed #475569;display:inline-block;border-radius:2px"></span> Sans tables définies
        </span>
      </div>

      <!-- Graph -->
      <div id="schema-cy" style="height:calc(100vh - 220px);background:var(--bg);border:1px solid var(--border);border-radius:8px"></div>

      <!-- Panneau détail (click) -->
      <div v-if="selected" style="position:absolute;top:48px;right:12px;width:240px;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px;font-size:0.82rem">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
          <strong>{{ selected.label }}</strong>
          <span style="cursor:pointer;color:var(--text-dim)" @click="selected=null">✕</span>
        </div>
        <div style="margin-bottom:6px">
          <span class="badge" :style="{background: funcColor(selected.owner_function).bg, color: funcColor(selected.owner_function).text}">
            {{ selected.owner_function }}
          </span>
        </div>
        <div style="color:var(--text-dim);margin-bottom:8px;font-size:0.78rem">{{ selected.description }}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:0.75rem;margin-bottom:10px">
          <div style="background:var(--surface2);padding:6px;border-radius:4px;text-align:center">
            <div style="font-size:1.1rem;font-weight:700">{{ selected.post_count }}</div>
            <div style="color:var(--text-dim)">Posts</div>
          </div>
          <div style="background:var(--surface2);padding:6px;border-radius:4px;text-align:center">
            <div style="font-size:1.1rem;font-weight:700">{{ selected.std_tables?.length||0 }}</div>
            <div style="color:var(--text-dim)">Tables</div>
          </div>
        </div>
        <button class="btn btn-primary btn-sm" style="width:100%" @click="$emit('open-class', selected.id)">
          Ouvrir la fiche →
        </button>
      </div>

      <!-- Wizard -->
      <class-wizard v-if="showWizard" :functions="functions" @save="saveClass" @cancel="showWizard=false" />
    </div>
  `,
  components: { ClassWizard }
};

// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Schema Classes (master-detail)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewSchemaClasses = {
  components: { TagsInput, TableStdEditor, ClassWizard },
  emits: ['toast'],
  props: ['initialClass'],
  setup(props, { emit }) {
    const classes = ref([]);
    const functions = ref([]);
    const selected = ref(null);
    const activeTab = ref('identite');
    const showWizard = ref(false);
    const editData = ref(null);
    const search = ref('');
    const mxlData = ref('');

    const filtered = computed(() =>
      classes.value.filter(c =>
        !search.value ||
        c.id.includes(search.value.toLowerCase()) ||
        c.label.toLowerCase().includes(search.value.toLowerCase())
      )
    );

    const load = async () => {
      [classes.value, functions.value] = await Promise.all([
        GET('/api/schema/classes'),
        GET('/api/schema/functions'),
      ]);
      if (props.initialClass) {
        const cls = classes.value.find(c => c.id === props.initialClass);
        if (cls) selectClass(cls);
      } else if (!selected.value && classes.value.length) {
        selectClass(classes.value[0]);
      }
    };
    onMounted(load);
    watch(() => props.initialClass, id => {
      if (id) {
        const cls = classes.value.find(c => c.id === id);
        if (cls) selectClass(cls);
      }
    });

    const selectClass = async (cls) => {
      selected.value = JSON.parse(JSON.stringify(cls));
      activeTab.value = 'identite';
      mxlData.value = '';
    };

    const loadMxl = async () => {
      if (!selected.value) return;
      const r = await GET(`/api/mxl/${selected.value.id}`).catch(() => null);
      mxlData.value = r?.mxl || '# Aucun Post de cette Classe n\'existe encore dans la Population';
    };
    watch(activeTab, t => { if (t === 'mxl') loadMxl(); });

    const openCreate = () => { editData.value = null; showWizard.value = true; };
    const openEdit = () => {
      if (!selected.value) return;
      editData.value = JSON.parse(JSON.stringify(selected.value));
      showWizard.value = true;
    };

    const saveWizard = async (form) => {
      try {
        if (editData.value) {
          await PUT(`/api/schema/classes/${form.id}`, form);
          emit('toast', { msg: 'Classe mise à jour', type: 'ok' });
        } else {
          await POST('/api/schema/classes', form);
          emit('toast', { msg: 'Classe créée', type: 'ok' });
        }
        showWizard.value = false;
        await load();
        const updated = classes.value.find(c => c.id === form.id);
        if (updated) selectClass(updated);
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const del = async () => {
      if (!selected.value) return;
      if (!confirm(`Supprimer la Classe "${selected.value.id}" ?`)) return;
      try {
        await DEL(`/api/schema/classes/${selected.value.id}`);
        emit('toast', { msg: 'Classe supprimée', type: 'ok' });
        selected.value = null;
        await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const fnLabel = id => functions.value.find(f => f.id === id)?.label || id;

    const mxlHighlighted = computed(() => {
      if (!mxlData.value) return '';
      return mxlData.value
        .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
        .replace(/^(FILE_TYPE|FILE_ID|VERSION)(\s+.*)$/gm,'<span style="color:#f97316;font-weight:600">$1</span><span>$2</span>')
        .replace(/^(DEF|COL|PULL|PUSH|COMPUTE|LIST|COLLECT)(\s+)/gm,'<span style="color:#60a5fa;font-weight:600">$1</span><span>$2</span>')
        .replace(/\b(TYPE|SHEET|HEADER|WRITE|KEY|FROM|TO|MODE|FORMULA)=/g,'<span style="color:#a78bfa">$1</span>=')
        .replace(/"([^"]*)"/g,'<span style="color:#34d399">"$1"</span>')
        .replace(/^(#.*)$/gm,'<span style="color:#475569;font-style:italic">$1</span>');
    });

    return {
      classes, functions, selected, activeTab, showWizard, editData,
      search, filtered, mxlData, mxlHighlighted,
      selectClass, openCreate, openEdit, saveWizard, del, fnLabel, funcColor,
    };
  },
  template: `
    <div style="display:flex;height:calc(100vh - 120px);gap:0">

      <!-- Panneau gauche : liste -->
      <div style="width:220px;flex-shrink:0;border-right:1px solid var(--border);display:flex;flex-direction:column">
        <div style="padding:10px">
          <input v-model="search" placeholder="Rechercher…" style="width:100%;font-size:0.82rem" />
        </div>
        <div style="flex:1;overflow-y:auto">
          <div v-for="cls in filtered" :key="cls.id"
               @click="selectClass(cls)"
               :style="selected?.id===cls.id ? 'background:var(--accent);' : ''"
               style="padding:10px 12px;cursor:pointer;border-bottom:1px solid var(--border)">
            <div style="font-size:0.85rem;font-weight:600;margin-bottom:2px"
                 :style="selected?.id===cls.id ? 'color:#fff' : ''">{{ cls.label }}</div>
            <div style="font-size:0.72rem" :style="selected?.id===cls.id ? 'color:#bfdbfe' : 'color:var(--text-dim)'">
              {{ cls.id }}
            </div>
            <div style="display:flex;gap:4px;margin-top:4px">
              <span class="badge" style="font-size:0.65rem"
                    :style="{background: funcColor(cls.owner_function).bg, color: funcColor(cls.owner_function).text}">
                {{ cls.owner_function }}
              </span>
              <span class="badge badge-gray" style="font-size:0.65rem">{{ cls.std_tables?.length||0 }} tables</span>
            </div>
          </div>
        </div>
        <div style="padding:10px;border-top:1px solid var(--border)">
          <button class="btn btn-primary btn-sm" style="width:100%" @click="openCreate">+ Nouvelle Classe</button>
        </div>
      </div>

      <!-- Panneau droit : fiche -->
      <div v-if="selected" style="flex:1;display:flex;flex-direction:column;overflow:hidden">

        <!-- En-tête fiche -->
        <div style="padding:16px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:12px">
          <div style="flex:1">
            <div style="font-size:1.1rem;font-weight:700">{{ selected.label }}</div>
            <div style="font-size:0.78rem;color:var(--text-dim);margin-top:2px">
              <code style="color:var(--accent)">{{ selected.id }}</code> ·
              <span :style="{color: funcColor(selected.owner_function).text}">{{ fnLabel(selected.owner_function) }}</span>
            </div>
          </div>
          <button class="btn btn-ghost btn-sm" @click="openEdit">✎ Modifier</button>
          <button class="btn btn-danger btn-sm" @click="del">Supprimer</button>
        </div>

        <!-- Onglets -->
        <div style="display:flex;border-bottom:1px solid var(--border);padding:0 20px">
          <div v-for="t in [{id:'identite',l:'Identité'},{id:'structure',l:'Structure'},{id:'champs',l:'Champs scalaires'},{id:'tables',l:'Tables std'},{id:'mxl',l:'Template MXL'}]"
               :key="t.id"
               @click="activeTab=t.id"
               style="padding:10px 14px;cursor:pointer;font-size:0.83rem;border-bottom:2px solid transparent;margin-bottom:-1px"
               :style="activeTab===t.id ? 'border-color:var(--accent);color:var(--text)' : 'color:var(--text-dim)'">
            {{ t.l }}
            <span v-if="t.id==='tables'" class="topbar-badge">{{ selected.std_tables?.length||0 }}</span>
            <span v-if="t.id==='champs'" class="topbar-badge">{{ selected.min_fields?.length||0 }}</span>
          </div>
        </div>

        <!-- Contenu onglet -->
        <div style="flex:1;overflow-y:auto;padding:20px">

          <!-- Identité -->
          <div v-if="activeTab==='identite'">
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px">
              <div>
                <div style="font-size:0.72rem;color:var(--text-dim);margin-bottom:4px">DESCRIPTION</div>
                <div style="font-size:0.85rem">{{ selected.description || '—' }}</div>
              </div>
              <div>
                <div style="font-size:0.72rem;color:var(--text-dim);margin-bottom:4px">FONCTION OWNER</div>
                <span class="badge" :style="{background: funcColor(selected.owner_function).bg, color: funcColor(selected.owner_function).text}">
                  {{ selected.owner_function }}
                </span>
                <div style="font-size:0.78rem;color:var(--text-dim);margin-top:4px">{{ fnLabel(selected.owner_function) }}</div>
              </div>
            </div>
            <div style="background:var(--surface2);border-radius:6px;padding:14px;font-size:0.82rem">
              <div style="color:var(--text-dim);margin-bottom:8px;font-size:0.72rem">RÈGLE DE MINIMUM</div>
              Les Posts de Classe <strong>{{ selected.id }}</strong> peuvent avoir plus de feuilles,
              de tables et de colonnes que ce qui est défini ici. Ce schéma définit le <em>contrat minimal</em>.
            </div>
          </div>

          <!-- Structure -->
          <div v-if="activeTab==='structure'">
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
              <div>
                <div style="font-size:0.72rem;color:var(--text-dim);margin-bottom:8px">FEUILLES MINIMALES</div>
                <div style="display:flex;flex-wrap:wrap;gap:6px">
                  <span v-for="s in selected.min_sheets" :key="s" class="badge badge-blue">{{ s }}</span>
                  <span v-if="!selected.min_sheets?.length" style="color:var(--text-dim);font-size:0.8rem">Aucune</span>
                </div>
              </div>
              <div>
                <div style="font-size:0.72rem;color:var(--text-dim);margin-bottom:8px">FEUILLES OPTIONNELLES</div>
                <div style="display:flex;flex-wrap:wrap;gap:6px">
                  <span v-for="s in selected.optional_sheets" :key="s" class="badge badge-gray">{{ s }}</span>
                  <span v-if="!selected.optional_sheets?.length" style="color:var(--text-dim);font-size:0.8rem">Aucune</span>
                </div>
              </div>
              <div>
                <div style="font-size:0.72rem;color:var(--text-dim);margin-bottom:8px">NAMESPACES STORE AUTORISÉS</div>
                <div style="display:flex;flex-wrap:wrap;gap:6px">
                  <span v-for="n in selected.allowed_namespaces" :key="n" class="badge badge-purple">{{ n }}</span>
                  <span v-if="!selected.allowed_namespaces?.length" style="color:var(--text-dim);font-size:0.8rem">Aucun</span>
                </div>
              </div>
              <div>
                <div style="font-size:0.72rem;color:var(--text-dim);margin-bottom:8px">PUSH PREFIX</div>
                <code style="color:var(--accent);font-size:0.85rem">{{ selected.push_prefix || '—' }}</code>
              </div>
            </div>
          </div>

          <!-- Champs scalaires -->
          <div v-if="activeTab==='champs'">
            <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:14px">
              Données unitaires (DEF) qui caractérisent chaque Post — identité, état opérationnel ou paramètre.
            </div>
            <div v-if="!selected.min_fields?.length" class="empty">
              Aucun champ scalaire défini<br>
              <button class="btn btn-ghost" style="margin-top:12px" @click="openEdit">+ Définir les champs</button>
            </div>
            <table v-else>
              <thead>
                <tr><th>Nom</th><th>Label</th><th>Type</th><th>Nature</th><th>Source</th><th>Requis</th><th>Push</th></tr>
              </thead>
              <tbody>
                <tr v-for="f in selected.min_fields" :key="f.name">
                  <td><code style="color:var(--accent)">{{ f.name }}</code></td>
                  <td>{{ f.label || '—' }}</td>
                  <td><span class="badge badge-gray">{{ f.field_type }}</span></td>
                  <td>
                    <span class="badge" :class="f.nature==='operationnel'?'badge-orange':f.nature==='parametre'?'badge-purple':'badge-blue'">
                      {{ f.nature }}
                    </span>
                  </td>
                  <td><span class="badge badge-gray" style="font-size:0.68rem">{{ f.source }}</span></td>
                  <td style="text-align:center"><span :style="f.required?'color:#22c55e':'color:var(--text-dim)'">{{ f.required ? '✓' : '○' }}</span></td>
                  <td style="text-align:center"><span :style="f.pushable?'color:#f97316':'color:var(--text-dim)'">{{ f.pushable ? '↑' : '—' }}</span></td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Tables standard -->
          <div v-if="activeTab==='tables'">
            <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:14px">
              Ces tables définissent le minimum attendu. Les Posts peuvent en avoir d'autres.
              Les colonnes marquées <strong>Min</strong> sont obligatoires dans tout Post.
            </div>
            <div v-if="!selected.std_tables?.length" class="empty">
              Aucune table standard définie<br>
              <button class="btn btn-ghost" style="margin-top:12px" @click="openEdit();activeTab='tables'">
                + Définir les tables
              </button>
            </div>
            <div v-for="tbl in selected.std_tables" :key="tbl.name" class="card" style="margin-bottom:10px">
              <div class="card-header">
                <span class="card-title">{{ tbl.name }} <span style="color:var(--text-dim);font-weight:400;font-size:0.78rem">SHEET={{ tbl.sheet }}</span></span>
                <span class="topbar-badge">{{ tbl.columns.length }} colonnes</span>
              </div>
              <table>
                <thead><tr><th>Colonne</th><th>Type</th><th>Header</th><th>Write</th><th>Clé</th><th>Min</th></tr></thead>
                <tbody>
                  <tr v-for="col in tbl.columns" :key="col.name">
                    <td><code style="color:var(--accent)">{{ col.name }}</code></td>
                    <td><span class="badge" :class="col.col_type==='KEY'?'badge-orange':'badge-gray'">{{ col.col_type }}</span></td>
                    <td style="color:var(--text-dim)">{{ col.header||'—' }}</td>
                    <td style="font-size:0.78rem;color:var(--text-dim)">{{ col.write||'—' }}</td>
                    <td style="text-align:center">{{ col.is_key ? '✓' : '' }}</td>
                    <td style="text-align:center">
                      <span :style="col.required ? 'color:#22c55e' : 'color:var(--text-dim)'">{{ col.required ? '✓' : '○' }}</span>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <!-- Template MXL -->
          <div v-if="activeTab==='mxl'">
            <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:12px">
              Généré depuis le premier Post de cette Classe dans la Population. Sert de base aux nouveaux Posts.
            </div>
            <pre v-if="mxlData" style="background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:14px;font-size:0.8rem;line-height:1.6;overflow-x:auto;font-family:monospace"
                 v-html="mxlHighlighted"></pre>
            <div v-else class="empty">Chargement…</div>
          </div>
        </div>
      </div>

      <div v-else style="flex:1;display:flex;align-items:center;justify-content:center;color:var(--text-dim);flex-direction:column;gap:12px">
        <div>← Sélectionner une Classe</div>
        <button class="btn btn-primary" @click="openCreate">+ Créer la première Classe</button>
      </div>

      <!-- Wizard -->
      <class-wizard v-if="showWizard" :functions="functions" :editData="editData"
                    @save="saveWizard" @cancel="showWizard=false" />
    </div>
  `
};

// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Schema Relations P/F
// ═══════════════════════════════════════════════════════════════════════════════
const ViewSchemaRelations = {
  emits: ['toast'],
  setup(_, { emit }) {
    const relations = ref([]);
    const classes = ref([]);
    const showModal = ref(false);
    const editing = ref(null);
    const blank = () => ({ id:'', parent_class:'', child_class:'', qualifier:'TYPICAL', cardinality:'1..N', description:'' });
    const form = reactive(blank());

    const load = async () => {
      [relations.value, classes.value] = await Promise.all([
        GET('/api/schema/relations'), GET('/api/schema/classes')
      ]);
    };
    onMounted(load);

    const classLabel = id => classes.value.find(c => c.id === id)?.label || id;

    const openCreate = () => { Object.assign(form, blank()); form.id = 'rel-'+Math.random().toString(36).slice(2,10); editing.value = null; showModal.value = true; };
    const openEdit = rel => { Object.assign(form, {...rel}); editing.value = rel.id; showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value) { await PUT(`/api/schema/relations/${editing.value}`, form); emit('toast', {msg:'Relation mise à jour',type:'ok'}); }
        else { await POST('/api/schema/relations', form); emit('toast', {msg:'Relation créée',type:'ok'}); }
        showModal.value = false; await load();
      } catch(e) { emit('toast', {msg:e.message,type:'error'}); }
    };

    const del = async id => {
      if (!confirm('Supprimer cette relation ?')) return;
      try { await DEL(`/api/schema/relations/${id}`); await load(); emit('toast', {msg:'Relation supprimée',type:'ok'}); }
      catch(e) { emit('toast', {msg:e.message,type:'error'}); }
    };

    return { relations, classes, showModal, editing, form, classLabel, openCreate, openEdit, save, del, funcColor };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <div>
            <span class="card-title">Relations Père / Fils <span class="topbar-badge">{{ relations.length }}</span></span>
            <div style="font-size:0.75rem;color:var(--text-dim);margin-top:2px">
              PRESCRIBED = contraignant (bloquant) · TYPICAL = indicatif (pattern attendu)
            </div>
          </div>
          <button class="btn btn-primary" @click="openCreate">+ Nouvelle relation</button>
        </div>

        <div v-if="!relations.length" class="empty">Aucune relation définie</div>
        <table v-else>
          <thead>
            <tr>
              <th>PÈRE</th><th></th><th>FILS</th><th>Qualifier</th><th>Cardinalité</th><th>Description</th><th></th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="r in relations" :key="r.id">
              <td>
                <div style="font-weight:600">{{ classLabel(r.parent_class) }}</div>
                <div style="font-size:0.72rem;color:var(--text-dim)">{{ r.parent_class }}</div>
              </td>
              <td style="text-align:center;color:var(--text-dim)">→</td>
              <td>
                <div style="font-weight:600">{{ classLabel(r.child_class) }}</div>
                <div style="font-size:0.72rem;color:var(--text-dim)">{{ r.child_class }}</div>
              </td>
              <td>
                <span class="badge" :class="r.qualifier==='PRESCRIBED' ? 'badge-orange' : 'badge-gray'"
                      style="cursor:pointer"
                      @click="r.qualifier=r.qualifier==='PRESCRIBED'?'TYPICAL':'PRESCRIBED';PUT('/api/schema/relations/'+r.id,r)">
                  {{ r.qualifier }}
                </span>
              </td>
              <td style="font-size:0.85rem">{{ r.cardinality }}</td>
              <td style="color:var(--text-dim);font-size:0.8rem">{{ r.description }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(r)">✎</button>
                <button class="btn btn-danger btn-sm" @click="del(r.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvelle' }} Relation P/F</div>
          <div class="form-row">
            <div class="form-group">
              <label>Classe Père *</label>
              <select v-model="form.parent_class">
                <option value="">— Choisir —</option>
                <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label }} ({{ c.id }})</option>
              </select>
            </div>
            <div class="form-group">
              <label>Classe Fils *</label>
              <select v-model="form.child_class">
                <option value="">— Choisir —</option>
                <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label }} ({{ c.id }})</option>
              </select>
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Qualifier</label>
              <select v-model="form.qualifier">
                <option value="TYPICAL">TYPICAL — indicatif</option>
                <option value="PRESCRIBED">PRESCRIBED — contraignant</option>
              </select>
            </div>
            <div class="form-group">
              <label>Cardinalité</label>
              <select v-model="form.cardinality">
                <option value="1..N">1..N (au moins un)</option>
                <option value="0..N">0..N (zéro ou plusieurs)</option>
                <option value="1..1">1..1 (exactement un)</option>
                <option value="0..1">0..1 (zéro ou un)</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <input v-model="form.description" placeholder="Ex: Un cockpit agrège les UO de son périmètre" />
          </div>
          <div style="background:var(--surface2);border-radius:6px;padding:10px 12px;font-size:0.8rem;color:var(--text-dim);margin-bottom:4px">
            <strong>{{ form.qualifier === 'PRESCRIBED' ? '⛔ PRESCRIBED' : '💡 TYPICAL' }}</strong> :
            {{ form.qualifier === 'PRESCRIBED'
               ? 'Tout Post de Classe ' + (form.child_class||'…') + ' devra avoir un Père de Classe ' + (form.parent_class||'…') + '.'
               : 'Un Post de Classe ' + (form.child_class||'…') + ' a généralement un Père de Classe ' + (form.parent_class||'…') + ', sans obligation.' }}
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
// VIEW: Schema Templates
// ═══════════════════════════════════════════════════════════════════════════════
const ViewSchemaTemplates = {
  components: { TagsInput, TableStdEditor },
  emits: ['toast'],
  setup(_, { emit }) {
    const templates = ref([]);
    const classes = ref([]);
    const selected = ref(null);
    const showModal = ref(false);
    const editing = ref(null);
    const filterClass = ref('');

    const blank = () => ({
      id: '', label: '', class_id: '', description: '',
      extra_sheets: [], field_defaults: {}, std_tables: [],
      mxl_defaults: {}, source_file: '',
    });
    const form = reactive(blank());

    const fieldDefaultsText = ref('{}');
    const mxlDefaultsText = ref('{}');

    const load = async () => {
      [templates.value, classes.value] = await Promise.all([
        GET('/api/schema/templates'), GET('/api/schema/classes')
      ]);
    };
    onMounted(load);

    const filtered = computed(() =>
      filterClass.value
        ? templates.value.filter(t => t.class_id === filterClass.value)
        : templates.value
    );

    const classLabel = id => classes.value.find(c => c.id === id)?.label || id;
    const classFields = id => classes.value.find(c => c.id === id)?.min_fields || [];

    const openCreate = () => {
      Object.assign(form, blank());
      fieldDefaultsText.value = '{}';
      mxlDefaultsText.value = '{}';
      editing.value = null;
      showModal.value = true;
    };

    const openEdit = tpl => {
      Object.assign(form, JSON.parse(JSON.stringify(tpl)));
      fieldDefaultsText.value = JSON.stringify(tpl.field_defaults || {}, null, 2);
      mxlDefaultsText.value = JSON.stringify(tpl.mxl_defaults || {}, null, 2);
      editing.value = tpl.id;
      showModal.value = true;
    };

    const save = async () => {
      try {
        form.field_defaults = JSON.parse(fieldDefaultsText.value || '{}');
        form.mxl_defaults   = JSON.parse(mxlDefaultsText.value || '{}');
      } catch { emit('toast', { msg: 'JSON invalide dans les defaults', type: 'error' }); return; }
      try {
        if (!form.id) form.id = 'tpl-' + Math.random().toString(36).slice(2, 10);
        if (editing.value) {
          await PUT(`/api/schema/templates/${editing.value}`, form);
          emit('toast', { msg: 'Template mis à jour', type: 'ok' });
        } else {
          await POST('/api/schema/templates', form);
          emit('toast', { msg: 'Template créé', type: 'ok' });
        }
        showModal.value = false;
        await load();
      } catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    const del = async id => {
      if (!confirm('Supprimer ce Template ?')) return;
      try { await DEL(`/api/schema/templates/${id}`); await load(); emit('toast', { msg: 'Template supprimé', type: 'ok' }); }
      catch(e) { emit('toast', { msg: e.message, type: 'error' }); }
    };

    return {
      templates, classes, selected, showModal, editing, filterClass,
      filtered, form, fieldDefaultsText, mxlDefaultsText,
      classLabel, classFields, openCreate, openEdit, save, del, funcColor,
    };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <div>
            <span class="card-title">Templates <span class="topbar-badge">{{ filtered.length }}</span></span>
            <div style="font-size:0.75rem;color:var(--text-dim);margin-top:2px">
              Réalisations concrètes d'une Classe — points de départ pour créer des Posts
            </div>
          </div>
          <div style="display:flex;gap:8px;align-items:center">
            <select v-model="filterClass" style="font-size:0.82rem;min-width:160px">
              <option value="">Toutes les Classes</option>
              <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label }}</option>
            </select>
            <button class="btn btn-primary" @click="openCreate">+ Nouveau Template</button>
          </div>
        </div>

        <div v-if="!filtered.length" class="empty">Aucun Template défini</div>
        <table v-else>
          <thead><tr><th>ID</th><th>Label</th><th>Classe</th><th>Feuilles extra</th><th>Tables</th><th>Fichier source</th><th></th></tr></thead>
          <tbody>
            <tr v-for="t in filtered" :key="t.id">
              <td><code style="color:var(--accent)">{{ t.id }}</code></td>
              <td><strong>{{ t.label }}</strong><br><span style="font-size:0.72rem;color:var(--text-dim)">{{ t.description }}</span></td>
              <td>
                <span class="badge badge-blue">{{ t.class_id }}</span><br>
                <span style="font-size:0.72rem;color:var(--text-dim)">{{ classLabel(t.class_id) }}</span>
              </td>
              <td style="font-size:0.78rem">{{ (t.extra_sheets||[]).join(', ') || '—' }}</td>
              <td style="text-align:center"><span class="topbar-badge">{{ t.std_tables?.length||0 }}</span></td>
              <td style="font-size:0.75rem;color:var(--text-dim)">{{ t.source_file || '—' }}</td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(t)">✎</button>
                <button class="btn btn-danger btn-sm" @click="del(t.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal" style="width:700px;max-width:98vw">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouveau' }} Template</div>

          <div class="form-row">
            <div class="form-group">
              <label>Classe parente *</label>
              <select v-model="form.class_id">
                <option value="">— Choisir —</option>
                <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label }} ({{ c.id }})</option>
              </select>
            </div>
            <div class="form-group">
              <label>ID Template</label>
              <input v-model="form.id" :disabled="!!editing" placeholder="tpl_uo_projet (auto si vide)" />
            </div>
          </div>
          <div class="form-group">
            <label>Label *</label>
            <input v-model="form.label" placeholder="UO Projet complexe" />
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="form.description" rows="2" placeholder="Pour UOs multi-acteurs avec suivi risque…" />
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Feuilles supplémentaires <span style="color:var(--text-dim)">(au-delà du minimum Classe)</span></label>
              <tags-input v-model="form.extra_sheets" placeholder="Dashboard, _Risques…" />
            </div>
            <div class="form-group">
              <label>Fichier Excel source</label>
              <input v-model="form.source_file" placeholder="templates/uo_projet.xlsx" />
            </div>
          </div>

          <div v-if="form.class_id && classFields(form.class_id).length" class="form-group">
            <label>Valeurs par défaut des champs scalaires</label>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
              <div v-for="f in classFields(form.class_id).filter(f=>f.source==='user_input')" :key="f.name">
                <label style="font-size:0.72rem">{{ f.label || f.name }}</label>
                <input :placeholder="f.name" @input="form.field_defaults[f.name]=$event.target.value" :value="form.field_defaults[f.name]||''" />
              </div>
            </div>
          </div>

          <div class="form-group">
            <label>Tables standard du Template <span style="color:var(--text-dim)">(surcharge / ajout vs Classe)</span></label>
            <div style="max-height:260px;overflow-y:auto">
              <table-std-editor v-model="form.std_tables" />
            </div>
          </div>

          <div class="form-group">
            <label>MXL defaults <span style="color:var(--text-dim)">(JSON — push_prefix, pull_keys…)</span></label>
            <textarea v-model="mxlDefaultsText" rows="3" style="font-family:monospace;font-size:0.8rem" placeholder='{"push_prefix":"uo.{id}.","pull_keys":["ref.acteurs"]}' />
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
// VIEW: Schema Functions
// ═══════════════════════════════════════════════════════════════════════════════
const ViewSchemaFunctions = {
  emits: ['toast'],
  setup(_, { emit }) {
    const functions = ref([]);
    const classes = ref([]);
    const showModal = ref(false);
    const editing = ref(null);
    const blank = () => ({ id:'', label:'', description:'', side:'interne' });
    const form = reactive(blank());

    const load = async () => {
      [functions.value, classes.value] = await Promise.all([
        GET('/api/schema/functions'), GET('/api/schema/classes')
      ]);
    };
    onMounted(load);

    const classesFor = fid => classes.value.filter(c => c.owner_function === fid).map(c => c.label);

    const openCreate = () => { Object.assign(form, blank()); editing.value = null; showModal.value = true; };
    const openEdit = fn => { Object.assign(form, {...fn}); editing.value = fn.id; showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value) { await PUT(`/api/schema/functions/${editing.value}`, form); emit('toast', {msg:'Fonction mise à jour',type:'ok'}); }
        else { await POST('/api/schema/functions', form); emit('toast', {msg:'Fonction créée',type:'ok'}); }
        showModal.value = false; await load();
      } catch(e) { emit('toast', {msg:e.message,type:'error'}); }
    };

    const del = async id => {
      if (!confirm('Supprimer cette Fonction ?')) return;
      try { await DEL(`/api/schema/functions/${id}`); await load(); emit('toast', {msg:'Fonction supprimée',type:'ok'}); }
      catch(e) { emit('toast', {msg:e.message,type:'error'}); }
    };

    return { functions, showModal, editing, form, classesFor, openCreate, openEdit, save, del, funcColor };
  },
  template: `
    <div>
      <div class="card">
        <div class="card-header">
          <div>
            <span class="card-title">Fonctions Acteurs <span class="topbar-badge">{{ functions.length }}</span></span>
            <div style="font-size:0.75rem;color:var(--text-dim);margin-top:2px">
              Les Fonctions définissent les rôles disponibles pour les Acteurs (humains).
            </div>
          </div>
          <button class="btn btn-primary" @click="openCreate">+ Nouvelle Fonction</button>
        </div>

        <table>
          <thead><tr><th>ID</th><th>Label</th><th>Side</th><th>Description</th><th>Classes utilisant</th><th></th></tr></thead>
          <tbody>
            <tr v-for="fn in functions" :key="fn.id">
              <td>
                <span class="badge" :style="{background: funcColor(fn.id).bg, color: funcColor(fn.id).text}">
                  {{ fn.id }}
                </span>
              </td>
              <td><strong>{{ fn.label }}</strong></td>
              <td>
                <span class="badge" :style="fn.side==='interne'
                  ? {background:'#1e3a5f',color:'#60a5fa'}
                  : {background:'#292524',color:'#d97706'}"
                  style="font-size:0.7rem">
                  {{ fn.side || 'interne' }}
                </span>
              </td>
              <td style="color:var(--text-dim);font-size:0.82rem">{{ fn.description }}</td>
              <td>
                <div style="display:flex;flex-wrap:wrap;gap:4px">
                  <span v-for="c in classesFor(fn.id)" :key="c" class="badge badge-gray" style="font-size:0.7rem">{{ c }}</span>
                  <span v-if="!classesFor(fn.id).length" style="color:var(--text-dim);font-size:0.78rem">—</span>
                </div>
              </td>
              <td style="display:flex;gap:6px;justify-content:flex-end">
                <button class="btn btn-ghost btn-sm" @click="openEdit(fn)">✎</button>
                <button class="btn btn-danger btn-sm" @click="del(fn.id)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal" style="width:480px">
          <div class="modal-title">{{ editing ? 'Modifier' : 'Nouvelle' }} Fonction</div>
          <div class="form-group">
            <label>ID unique *</label>
            <input v-model="form.id" :disabled="!!editing" placeholder="ingenieur_sys" />
          </div>
          <div class="form-group">
            <label>Label *</label>
            <input v-model="form.label" placeholder="Ingénieur Systèmes" />
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="form.description" rows="2" placeholder="Responsabilités de cette Fonction…" />
          </div>
          <div class="form-group">
            <label>Side *</label>
            <select v-model="form.side">
              <option value="interne">interne — Acteur interne à l'organisation</option>
              <option value="externe">externe — Acteur externe (client, expert, fournisseur…)</option>
            </select>
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
