// ═══════════════════════════════════════════════════════════════════════════════
// SCHEMA VIEWS — Composants partagés N1 Schema Designer
// ═══════════════════════════════════════════════════════════════════════════════

// ── Constantes ────────────────────────────────────────────────────────────────

const COL_TYPES   = ['string', 'int', 'float', 'date', 'bool', 'ref', 'pct', 'KEY'];
const WRITE_MODES = ['', 'ingenieur_sys', 'pilote_metier', 'pilote_pole', 'system', 'user'];
const FIELD_NATURES  = ['identitaire', 'operationnel', 'parametre'];
const FIELD_SOURCES  = ['user_input', 'system', 'computed', 'reference'];
const FIELD_TYPES    = ['string', 'int', 'float', 'date', 'bool', 'ref'];
const QUALIFIERS     = ['PRESCRIBED', 'TYPICAL'];

// Couleurs par fonction owner (pour le Blueprint)
const FUNC_COLORS = {
  pilote_pole:          { bg: '#431407', border: '#f97316', text: '#fb923c' },
  pilote_metier:        { bg: '#14532d', border: '#22c55e', text: '#4ade80' },
  ingenieur_sys:        { bg: '#1e3a5f', border: '#3b82f6', text: '#60a5fa' },
  it_manager:           { bg: '#3b0764', border: '#a855f7', text: '#c084fc' },
  donneur_ordre_pole:   { bg: '#292524', border: '#d97706', text: '#fbbf24' },
  donneur_ordre_metier: { bg: '#1c1917', border: '#78716c', text: '#a8a29e' },
  ingenieur_sys_client: { bg: '#134e4a', border: '#14b8a6', text: '#2dd4bf' },
  expert_technique:     { bg: '#4a044e', border: '#ec4899', text: '#f472b6' },
  fournisseur:          { bg: '#1e293b', border: '#64748b', text: '#94a3b8' },
};
const DEFAULT_FUNC_COLOR = { bg: '#1e293b', border: '#6366f1', text: '#818cf8' };
const funcColor = id => FUNC_COLORS[id] || DEFAULT_FUNC_COLOR;


// ── TagsInput ─────────────────────────────────────────────────────────────────

const TagsInput = {
  props: ['modelValue', 'placeholder'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const input = Vue.ref('');
    const tags = Vue.computed(() => props.modelValue || []);
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
    <div class="tags-input" @click="$el.querySelector('input').focus()">
      <span v-for="t in tags" :key="t" class="tag">
        {{ t }} <span class="tag-remove" @click.stop="remove(t)">×</span>
      </span>
      <input v-model="input" :placeholder="placeholder || 'Ajouter...'" @keydown="onKey" @blur="add" />
    </div>
  `
};


// ── ColEditor ─────────────────────────────────────────────────────────────────

const ColEditor = {
  props: ['modelValue'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const cols = Vue.computed(() => props.modelValue || []);
    const newCol = () => ({ name:'', col_type:'string', header:'', write:'', is_key:false, required:true, description:'' });
    const add    = () => emit('update:modelValue', [...cols.value, newCol()]);
    const remove = i => emit('update:modelValue', cols.value.filter((_,j) => j!==i));
    const upd    = (i, k, v) => emit('update:modelValue', cols.value.map((c,j) => j===i ? {...c,[k]:v} : c));
    return { cols, add, remove, upd, COL_TYPES, WRITE_MODES };
  },
  template: `
    <div>
      <div style="display:flex;justify-content:flex-end;margin-bottom:6px">
        <button class="btn btn-ghost btn-sm" @click="add">+ Colonne</button>
      </div>
      <table v-if="cols.length">
        <thead>
          <tr><th>Nom</th><th>Type</th><th>Header</th><th>Write</th><th title="Clé">Clé</th><th title="Requis">Min</th><th></th></tr>
        </thead>
        <tbody>
          <tr v-for="(col,i) in cols" :key="i">
            <td><input :value="col.name"     @input="upd(i,'name',$event.target.value)"     placeholder="nom_col" /></td>
            <td>
              <select :value="col.col_type" @change="upd(i,'col_type',$event.target.value)" style="min-width:72px">
                <option v-for="t in COL_TYPES" :key="t" :value="t">{{ t }}</option>
              </select>
            </td>
            <td><input :value="col.header"   @input="upd(i,'header',$event.target.value)"   placeholder="Libellé" /></td>
            <td>
              <select :value="col.write" @change="upd(i,'write',$event.target.value)" style="min-width:80px">
                <option v-for="w in WRITE_MODES" :key="w" :value="w">{{ w||'—' }}</option>
              </select>
            </td>
            <td style="text-align:center"><input type="checkbox" :checked="col.is_key"  @change="upd(i,'is_key',$event.target.checked)"  style="width:auto;accent-color:var(--accent)" /></td>
            <td style="text-align:center"><input type="checkbox" :checked="col.required" @change="upd(i,'required',$event.target.checked)" style="width:auto;accent-color:var(--accent)" /></td>
            <td><button class="btn btn-danger btn-xs" @click="remove(i)">✕</button></td>
          </tr>
        </tbody>
      </table>
      <div v-else style="color:var(--text-dim);font-size:0.78rem;padding:8px 0">Aucune colonne — cliquez "+ Colonne"</div>
    </div>
  `
};


// ── TableStdEditor ────────────────────────────────────────────────────────────

const TableStdEditor = {
  components: { ColEditor },
  props: ['modelValue'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const tables   = Vue.computed(() => props.modelValue || []);
    const expanded = Vue.ref(null);
    const newTable = () => ({ name:'', sheet:'', columns:[], description:'' });
    const add    = () => { emit('update:modelValue', [...tables.value, newTable()]); expanded.value = tables.value.length; };
    const remove = i => { if (expanded.value === i) expanded.value = null; emit('update:modelValue', tables.value.filter((_,j)=>j!==i)); };
    const upd    = (i, k, v) => emit('update:modelValue', tables.value.map((t,j) => j===i ? {...t,[k]:v} : t));
    const toggle = i => { expanded.value = expanded.value === i ? null : i; };
    return { tables, expanded, add, remove, upd, toggle };
  },
  template: `
    <div>
      <div style="display:flex;justify-content:flex-end;margin-bottom:8px">
        <button class="btn btn-ghost btn-sm" @click="add">+ Table</button>
      </div>
      <div v-for="(tbl,i) in tables" :key="i"
           style="border:1px solid var(--border);border-radius:7px;margin-bottom:8px;overflow:hidden">
        <div style="display:flex;align-items:center;gap:8px;padding:9px 12px;cursor:pointer;background:var(--surface2)" @click="toggle(i)">
          <span style="flex:1;font-size:0.83rem;font-weight:600">
            {{ tbl.name || '(sans nom)' }}
            <span style="color:var(--text-dim);font-weight:400;font-size:0.72rem"> — {{ tbl.sheet||'feuille?' }}</span>
            <span class="badge badge-gray" style="margin-left:6px">{{ tbl.columns.length }} col.</span>
          </span>
          <span style="color:var(--text-dim);font-size:0.8rem">{{ expanded===i ? '▲' : '▼' }}</span>
          <button class="btn btn-danger btn-xs" @click.stop="remove(i)">✕</button>
        </div>
        <div v-if="expanded===i" style="padding:12px;border-top:1px solid var(--border)">
          <div class="form-row" style="margin-bottom:10px">
            <div class="form-group" style="margin-bottom:0">
              <label>Nom table (schéma)</label>
              <input :value="tbl.name" @input="upd(i,'name',$event.target.value)" placeholder="TabActivites" />
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label>Feuille Excel</label>
              <input :value="tbl.sheet" @input="upd(i,'sheet',$event.target.value)" placeholder="Activites" />
            </div>
          </div>
          <col-editor :modelValue="tbl.columns" @update:modelValue="upd(i,'columns',$event)" />
        </div>
      </div>
      <div v-if="!tables.length" style="color:var(--text-dim);font-size:0.78rem;padding:8px 0">
        Aucune table standard
      </div>
    </div>
  `
};


// ── FieldEditor ───────────────────────────────────────────────────────────────

const FieldEditor = {
  props: ['modelValue'],
  emits: ['update:modelValue'],
  setup(props, { emit }) {
    const fields   = Vue.computed(() => props.modelValue || []);
    const expanded = Vue.ref(null);
    const newField = () => ({ name:'', label:'', field_type:'string', nature:'identitaire', source:'user_input', required:true, pushable:false, ref_class:'', description:'' });
    const add    = () => { emit('update:modelValue', [...fields.value, newField()]); expanded.value = fields.value.length; };
    const remove = i => { if (expanded.value===i) expanded.value=null; emit('update:modelValue', fields.value.filter((_,j)=>j!==i)); };
    const upd    = (i, k, v) => emit('update:modelValue', fields.value.map((f,j)=>j===i?{...f,[k]:v}:f));
    const toggle = i => { expanded.value = expanded.value===i ? null : i; };
    return { fields, expanded, add, remove, upd, toggle, FIELD_NATURES, FIELD_SOURCES, FIELD_TYPES };
  },
  template: `
    <div>
      <div style="display:flex;justify-content:flex-end;margin-bottom:8px">
        <button class="btn btn-ghost btn-sm" @click="add">+ Champ</button>
      </div>
      <div v-for="(f,i) in fields" :key="i"
           style="border:1px solid var(--border);border-radius:7px;margin-bottom:6px;overflow:hidden">
        <div style="display:flex;align-items:center;gap:8px;padding:8px 12px;cursor:pointer;background:var(--surface2)" @click="toggle(i)">
          <span style="flex:1;font-size:0.83rem;font-weight:500">
            {{ f.name || '(sans nom)' }}
            <span style="color:var(--text-dim);font-weight:400;font-size:0.7rem"> {{ f.field_type }}</span>
            <span class="badge badge-gray" style="margin-left:4px">{{ f.nature }}</span>
            <span v-if="f.pushable" style="color:var(--ok);font-size:0.7rem;margin-left:4px">↑PUSH</span>
          </span>
          <span style="color:var(--text-dim);font-size:0.8rem">{{ expanded===i ? '▲' : '▼' }}</span>
          <button class="btn btn-danger btn-xs" @click.stop="remove(i)">✕</button>
        </div>
        <div v-if="expanded===i" style="padding:12px;border-top:1px solid var(--border)">
          <div class="form-row" style="margin-bottom:10px">
            <div class="form-group" style="margin-bottom:0">
              <label>Identifiant (name)</label>
              <input :value="f.name"  @input="upd(i,'name',$event.target.value)"  placeholder="project_id" />
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label>Libellé (label)</label>
              <input :value="f.label" @input="upd(i,'label',$event.target.value)" placeholder="Identifiant projet" />
            </div>
          </div>
          <div class="form-row-3" style="margin-bottom:10px">
            <div class="form-group" style="margin-bottom:0">
              <label>Type</label>
              <select :value="f.field_type" @change="upd(i,'field_type',$event.target.value)">
                <option v-for="t in FIELD_TYPES" :key="t" :value="t">{{ t }}</option>
              </select>
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label>Nature</label>
              <select :value="f.nature" @change="upd(i,'nature',$event.target.value)">
                <option v-for="n in FIELD_NATURES" :key="n" :value="n">{{ n }}</option>
              </select>
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label>Source</label>
              <select :value="f.source" @change="upd(i,'source',$event.target.value)">
                <option v-for="s in FIELD_SOURCES" :key="s" :value="s">{{ s }}</option>
              </select>
            </div>
          </div>
          <div style="display:flex;gap:20px;font-size:0.83rem">
            <label class="toggle"><input type="checkbox" :checked="f.required"  @change="upd(i,'required',$event.target.checked)"  /> Requis</label>
            <label class="toggle"><input type="checkbox" :checked="f.pushable"  @change="upd(i,'pushable',$event.target.checked)"  /> Pushable</label>
          </div>
        </div>
      </div>
      <div v-if="!fields.length" style="color:var(--text-dim);font-size:0.78rem;padding:8px 0">
        Aucun champ scalaire
      </div>
    </div>
  `
};


// ── MxlHighlight ──────────────────────────────────────────────────────────────

const MxlHighlight = {
  props: ['code'],
  computed: {
    html() {
      if (!this.code) return '';
      return this.code
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/^(#.*)$/gm, '<span class="mxl-comment">$1</span>')
        .replace(/\b(FILE_TYPE|FILE_ID|VERSION|DEF|COL|PUSH|PULL|LIST|COLLECT|GET_TABLE|GET_CELL|FILL_TABLE|FORM|SOURCE|FILTER|WHERE|MODE|INTO|FROM_LIST|COLS|WITH|KEY|WRITE|HEADER|TYPE)\b/g,
                 '<span class="mxl-keyword">$1</span>')
        .replace(/"([^"]+)"/g, '<span class="mxl-string">"$1"</span>')
        .replace(/(\$\w+)/g, '<span class="mxl-value">$1</span>');
    }
  },
  template: `<div class="mxl-preview" v-html="html"></div>`
};
