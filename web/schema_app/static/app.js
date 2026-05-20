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
const GET  = path        => api('GET',    path);
const POST = (path, b)   => api('POST',   path, b);
const PUT  = (path, b)   => api('PUT',    path, b);
const DEL  = path        => api('DELETE', path);

// ── Toast ─────────────────────────────────────────────────────────────────────

const toasts = ref([]);
let toastId = 0;
function toast(msg, type = 'ok') {
  const id = ++toastId;
  toasts.value.push({ id, msg, type });
  setTimeout(() => { toasts.value = toasts.value.filter(t => t.id !== id); }, 3000);
}
function toastErr(e) { toast(e instanceof Error ? e.message : String(e), 'error'); }

const ToastLayer = {
  props: ['toasts'],
  template: `
    <div style="position:fixed;bottom:24px;right:24px;z-index:999;display:flex;flex-direction:column;gap:8px;">
      <transition-group name="toast">
        <div v-for="t in toasts" :key="t.id"
          :style="{background: t.type==='error'?'#7f1d1d':'#14532d',
                   border:'1px solid '+(t.type==='error'?'#ef4444':'#22c55e'),
                   padding:'10px 16px',borderRadius:'8px',fontSize:'0.83rem',
                   color:'#fff',maxWidth:'340px',boxShadow:'0 4px 12px rgba(0,0,0,0.5)'}">
          {{ t.msg }}
        </div>
      </transition-group>
    </div>`
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Gestionnaire d'Écosystèmes
// ═══════════════════════════════════════════════════════════════════════════════

const ViewEcosystemManager = {
  emits: ['open'],
  setup(_, { emit }) {
    const ecosystems = ref([]);
    const showNew = ref(false);
    const showOpen = ref(false);
    const newName = ref('');
    const openPath = ref('');

    const load = async () => {
      ecosystems.value = await GET('/api/ecosystem/list').catch(() => []);
    };

    onMounted(load);

    const openEco = async eco => {
      try {
        await POST('/api/ecosystem/open', { path: eco.path });
        toast(`Écosystème "${eco.name}" ouvert`);
        emit('open', eco.name);
      } catch(e) { toastErr(e); }
    };

    const createEco = async () => {
      try {
        await POST('/api/ecosystem/new', { name: newName.value });
        toast(`Écosystème "${newName.value}" créé`);
        showNew.value = false;
        newName.value = '';
        await load();
        emit('open', newName.value);
      } catch(e) { toastErr(e); }
    };

    const doOpen = async () => {
      try {
        const eco = await POST('/api/ecosystem/open', { path: openPath.value });
        toast(`Écosystème ouvert`);
        showOpen.value = false;
        openPath.value = '';
        await load();
        emit('open', eco.name);
      } catch(e) { toastErr(e); }
    };

    const deleteEco = async eco => {
      if (!confirm(`Retirer "${eco.name}" de la liste ?\n(les fichiers ne seront PAS supprimés)`)) return;
      await DEL(`/api/ecosystem/${encodeURIComponent(eco.name)}`).catch(toastErr);
      await load();
    };

    return { ecosystems, showNew, showOpen, newName, openPath, openEco, createEco, doOpen, deleteEco };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:20px">
        <span class="card-title" style="font-size:1.05rem">Choisir un écosystème</span>
        <div style="display:flex;gap:10px">
          <button class="btn btn-ghost" @click="showOpen=true">📂 Ouvrir</button>
          <button class="btn btn-primary" @click="showNew=true">+ Nouvel écosystème</button>
        </div>
      </div>
      <p style="color:var(--text-dim);font-size:0.83rem;margin-bottom:4px">
        Un écosystème est un dossier de configuration (file_types.yaml, relations, fonctions, templates).
      </p>
      <div class="eco-grid">
        <div v-for="eco in ecosystems" :key="eco.path"
             class="eco-card" :class="{active:eco.active}"
             @click="openEco(eco)">
          <div style="display:flex;align-items:start;justify-content:space-between;margin-bottom:6px">
            <div class="eco-card-name">
              {{ eco.name }}
              <span v-if="eco.active" class="badge badge-indigo" style="margin-left:8px;font-size:0.65rem">actif</span>
            </div>
            <button v-if="!eco.active" class="btn btn-danger btn-xs" @click.stop="deleteEco(eco)" title="Retirer">✕</button>
          </div>
          <div class="eco-card-path">{{ eco.path }}</div>
          <div class="eco-card-meta">
            <span>{{ eco.class_count }} classe{{ eco.class_count!==1?'s':'' }}</span>
            <span :style="{color: eco.valid ? 'var(--ok)' : 'var(--err)'}">{{ eco.valid ? '✓ valide' : '⚠ invalide' }}</span>
          </div>
        </div>
      </div>

      <!-- Modal : Nouvel écosystème -->
      <div v-if="showNew" class="modal-overlay" @click.self="showNew=false">
        <div class="modal">
          <div class="modal-title">Nouvel écosystème</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px;">
            Un sous-dossier sera créé automatiquement sous <code>gabarits_dir/&lt;nom&gt;/</code>.
          </div>
          <div class="form-group">
            <label>Nom *</label>
            <input v-model="newName" placeholder="SysEng-v2" autofocus />
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showNew=false">Annuler</button>
            <button class="btn btn-primary" @click="createEco" :disabled="!newName">Créer</button>
          </div>
        </div>
      </div>

      <!-- Modal : Ouvrir -->
      <div v-if="showOpen" class="modal-overlay" @click.self="showOpen=false">
        <div class="modal">
          <div class="modal-title">Ouvrir un écosystème existant</div>
          <div class="form-group">
            <label>Chemin du dossier config (doit contenir file_types.yaml)</label>
            <input v-model="openPath" placeholder="C:/chemin/vers/config" autofocus />
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showOpen=false">Annuler</button>
            <button class="btn btn-primary" @click="doOpen" :disabled="!openPath">Ouvrir</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Blueprint (graphe Cytoscape)
// ═══════════════════════════════════════════════════════════════════════════════

const ViewBlueprint = {
  emits: ['edit-class'],
  setup(_, { emit }) {
    const loading   = ref(true);
    const panel     = ref(null);   // nœud (Classe) cliqué
    const edgePanel = ref(null);   // arête (Relation) cliquée
    const allClasses = ref([]);    // cache pour les tables par classe
    let cy = null;

    // ── Flux helpers ──────────────────────────────────────────────────────────
    const customChild  = ref('');
    const customParent = ref('');

    async function openEdgePanel(rel) {
      try {
        const [full, parent, child] = await Promise.all([
          GET(`/api/relations/${rel.id}`),
          GET(`/api/classes/${rel.source}`).catch(() => null),
          GET(`/api/classes/${rel.target}`).catch(() => null),
        ]);
        edgePanel.value = {
          rel: full,
          parent,                                                       // données fraîches (re-fetchées à chaque ouverture)
          child,                                                        // idem — namespace/min_fields toujours à jour
          parentId:     rel.source,                                    // consommateur (dashboard)
          childId:      rel.target,                                    // source de données (posts)
          childTables:  child?.std_tables?.map(t => t.name)  ?? [],   // tables de l'enfant
          parentTables: parent?.std_tables?.map(t => t.name) ?? [],   // tables du parent
          isOneToMany:  (full.cardinality || '1..N').includes('N'),   // PULL interdit si 1..N
        };
        panel.value = null;
        customChild.value  = '';
        customParent.value = '';
      } catch(e) { console.error(e); }
    }

    function getFluxEntry(table, source, mode) {
      return edgePanel.value?.rel?.flux?.find(
        f => f.table === table && (f.source || 'child') === source && f.mode === mode
      ) ?? null;
    }

    function isActive(table, source, mode) {
      return !!getFluxEntry(table, source, mode);
    }

    // Auto-crée une table (std_tables) ou un scalaire (min_fields) dans une Classe si absent.
    // sourceColumns : colonnes à recopier (namespace + data). forceTable : forcer std_tables même pour un scalaire.
    // Retourne true si création effectuée, false si déjà présent.
    async function ensureInClass(classId, name, sourceColumns = [], forceTable = false) {
      const isTable = name.startsWith('Tab') || forceTable;
      try {
        const cls = await GET(`/api/classes/${classId}`);
        if (isTable) {
          if (cls.std_tables?.some(t => t.name === name)) return false;
          await PUT(`/api/classes/${classId}`, {
            ...cls,
            std_tables: [...(cls.std_tables || []), {
              name, sheet: name,
              columns: sourceColumns,   // structure copiée depuis la classe source
            }],
          });
        } else {
          if (cls.min_fields?.some(f => f.name === name)) return false;
          await PUT(`/api/classes/${classId}`, {
            ...cls,
            min_fields: [...(cls.min_fields || []), {
              name, label: name, field_type: 'string',
              nature: 'operationnel', source: 'user_input',
              required: false, pushable: false,
            }],
          });
        }
        return true;
      } catch(e) {
        console.error('ensureInClass:', e);
        return false;
      }
    }

    async function toggleFluxEntry(table, source, mode) {
      if (!table) return;
      const ep = edgePanel.value;
      if (!ep) return;
      const existing = getFluxEntry(table, source, mode);
      if (existing) {
        await DEL(`/api/relations/${ep.rel.id}/flux/${existing.id}`);
        ep.rel.flux = ep.rel.flux.filter(f => f.id !== existing.id);
        toast('Flux retiré');
      } else {
        // 1. Déclarer le flux
        const created = await POST(`/api/relations/${ep.rel.id}/flux`, { table, source, mode });
        ep.rel.flux.push(created);
        if (source === 'child')  customChild.value  = '';
        if (source === 'parent') customParent.value = '';

        // 2. Récupérer colonnes source + colonnes namespace si COLLECT
        //    PULL/COLLECT : source = child  ⟹  dest = parent
        //    PUSH         : source = parent ⟹  dest = child
        const destId   = (source === 'child') ? ep.parentId : ep.childId;
        // Utiliser les données fraîches stockées dans edgePanel (fetched à l'ouverture)
        const srcClass = (source === 'child') ? ep.child : ep.parent;
        const srcTable = srcClass?.std_tables?.find(t => t.name === table);
        const srcCols  = srcTable?.columns ?? [];

        // Colonnes namespace (COLLECT uniquement) = min_fields identitaires du child
        // Chaque ligne collectée doit porter l'identité complète du Post source
        let nsCols = [];
        if (mode === 'COLLECT') {
          nsCols = (ep.child?.min_fields ?? [])
            .filter(f => f.nature === 'identitaire')
            .map(f => ({
              name: f.name, col_type: f.field_type || 'string',
              header: f.label || f.name, is_key: true,
              required: true, write: '', description: `namespace: ${f.label || f.name}`,
            }));
        }

        // Pour un scalaire collecté (non-Tab) → forcer création en std_table
        // avec colonnes namespace + colonne valeur (le scalaire devient une ligne par Post)
        const isScalarCollect = mode === 'COLLECT' && !table.startsWith('Tab');
        const forceTable = isScalarCollect;
        const scalarValueCol = isScalarCollect ? [{
          name: table, col_type: srcClass?.min_fields?.find(f => f.name === table)?.field_type || 'string',
          header: table, is_key: false, required: false, write: '', description: '',
        }] : [];
        const finalCols = [...nsCols, ...srcCols, ...scalarValueCol];

        // 3. Auto-créer dans la Classe réceptrice
        const wasCreated = await ensureInClass(destId, table, finalCols, forceTable);
        if (wasCreated) {
          const palette = (source === 'child') ? ep.parentTables : ep.childTables;
          if (!palette.includes(table)) palette.push(table);
          const nsInfo  = nsCols.length  ? ` +${nsCols.length} ns`  : '';
          const colInfo = srcCols.length ? ` ${srcCols.length} col.` : '';
          toast(`"${table}" créé dans ${destId}${nsInfo}${colInfo}`);
        } else {
          toast('Flux ajouté');
        }
      }
    }

    async function removeFlux(fluxId) {
      const ep = edgePanel.value;
      if (!ep) return;
      await DEL(`/api/relations/${ep.rel.id}/flux/${fluxId}`);
      ep.rel.flux = ep.rel.flux.filter(f => f.id !== fluxId);
      toast('Flux supprimé');
    }

    onMounted(async () => {
      const [data, classes] = await Promise.all([
        GET('/api/blueprint').catch(() => ({ classes: [], relations: [] })),
        GET('/api/classes').catch(() => []),
      ]);
      allClasses.value = classes;
      loading.value = false;

      await nextTick();
      const elements = [];

      for (const cls of data.classes) {
        const c = funcColor(cls.owner_function);
        elements.push({
          data: { id: cls.id, label: cls.label || cls.id, owner: cls.owner_function, ...cls },
          style: {
            'background-color': c.bg,
            'border-color': c.border,
            'color': c.text,
          }
        });
      }
      for (const rel of data.relations) {
        elements.push({
          data: {
            id: rel.id, source: rel.parent_class, target: rel.child_class,
            qualifier: rel.qualifier, label: rel.qualifier,
          }
        });
      }

      cy = cytoscape({
        container: document.getElementById('cy-blueprint'),
        elements,
        style: [
          { selector: 'node', style: {
              'label': 'data(label)', 'text-valign': 'center', 'text-halign': 'center',
              'font-size': '11px', 'font-weight': '600', 'border-width': 2,
              'width': 120, 'height': 44, 'shape': 'round-rectangle',
              'text-wrap': 'wrap', 'text-max-width': '110px',
          }},
          { selector: 'edge', style: {
              'width': 1.5, 'target-arrow-shape': 'triangle', 'curve-style': 'bezier',
              'line-color': 'var(--border)', 'target-arrow-color': 'var(--border)',
              'font-size': '9px', 'label': 'data(label)', 'color': '#64748b',
              'text-background-color': '#0f172a', 'text-background-opacity': 1,
              'text-background-padding': '2px',
          }},
          { selector: 'edge[qualifier="PRESCRIBED"]', style: {
              'line-color': '#ef4444', 'target-arrow-color': '#ef4444',
              'line-style': 'solid', 'width': 2, 'color': '#fca5a5',
          }},
          { selector: 'edge[qualifier="TYPICAL"]', style: {
              'line-color': '#475569', 'target-arrow-color': '#475569',
              'line-style': 'dashed', 'color': '#64748b',
          }},
          { selector: ':selected', style: { 'border-color': '#818cf8', 'border-width': 3 }},
        ],
        layout: { name: 'breadthfirst', directed: true, spacingFactor: 1.4, padding: 40 },
        userZoomingEnabled: true,
        userPanningEnabled: true,
      });

      cy.on('tap', 'node', e => {
        panel.value = e.target.data();
        edgePanel.value = null;
      });
      cy.on('dbltap', 'node', e => {
        emit('edit-class', e.target.data().id);
      });
      cy.on('tap', 'edge', e => {
        panel.value = null;
        openEdgePanel(e.target.data());
      });
      cy.on('tap', e => {
        if (e.target === cy) { panel.value = null; edgePanel.value = null; }
      });
    });

    const relayout = () => cy && cy.layout({ name: 'cose', padding: 40 }).run();

    return {
      loading, panel, edgePanel, relayout, emit,
      customChild, customParent,
      isActive, toggleFluxEntry, removeFlux,
    };
  },
  template: `
    <div style="position:relative">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:12px">
        <span style="font-size:0.78rem;color:var(--text-dim)">
          Cliquez un nœud · Double-cliquez pour éditer · Cliquez une arête pour définir les flux
        </span>
        <button class="btn btn-ghost btn-sm" @click="relayout">⟳ Réorganiser</button>
        <span style="margin-left:auto;display:flex;gap:10px;font-size:0.75rem;color:var(--text-dim)">
          <span style="color:#fca5a5">— PRESCRIBED</span>
          <span style="color:#64748b">--- TYPICAL</span>
        </span>
      </div>
      <div v-if="loading" class="loading">Chargement du Blueprint…</div>
      <div style="position:relative" v-show="!loading">
        <div id="cy-blueprint"></div>

        <!-- Panneau latéral : Classe (nœud) -->
        <div v-if="panel" style="position:absolute;top:0;right:0;width:260px;background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px;z-index:10">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
            <span style="font-weight:600;font-size:0.9rem">{{ panel.label }}</span>
            <button class="btn btn-ghost btn-xs" @click="panel=null">✕</button>
          </div>
          <div style="font-size:0.78rem;color:var(--text-dim);margin-bottom:6px">{{ panel.id }}</div>
          <div v-if="panel.owner" style="margin-bottom:8px">
            <span class="badge badge-gray">{{ panel.owner }}</span>
          </div>
          <div v-if="panel.description" style="font-size:0.78rem;color:var(--text-dim);margin-bottom:12px">{{ panel.description }}</div>
          <div v-if="panel.min_sheets && panel.min_sheets.length" style="font-size:0.75rem;color:var(--text-dim);margin-bottom:12px">
            Feuilles : {{ panel.min_sheets.join(', ') }}
          </div>
          <button class="btn btn-primary btn-sm" style="width:100%" @click="emit('edit-class', panel.id)">
            ✏ Éditer cette Classe
          </button>
        </div>

      </div>

      <!-- Modal Flux de Schéma (arête cliquée) -->
      <div v-if="edgePanel" class="modal-overlay" @click.self="edgePanel=null">
        <div class="modal" style="max-width:740px;width:95%;">
          <!-- En-tête -->
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px;">
            <div>
              <div style="font-weight:700;font-size:1rem;margin-bottom:6px;">Flux de Schéma N1</div>
              <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
                <span class="badge badge-teal">{{ edgePanel.parentId }}</span>
                <span style="color:var(--text-dim);">→</span>
                <span class="badge badge-teal">{{ edgePanel.childId }}</span>
                <span class="badge" :class="edgePanel.rel.qualifier==='PRESCRIBED'?'badge-red':'badge-gray'">
                  {{ edgePanel.rel.qualifier }}
                </span>
                <span style="font-size:0.75rem;color:var(--text-dim);">{{ edgePanel.rel.cardinality }}</span>
              </div>
            </div>
            <button class="btn btn-ghost btn-sm" @click="edgePanel=null">✕</button>
          </div>

          <!-- Deux colonnes -->
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">

            <!-- Colonne gauche : deux palettes (child + parent) -->
            <div>

              <!-- Section CHILD : source de données (Posts) -->
              <div style="background:var(--surface2);border-radius:8px;padding:12px;margin-bottom:10px;">
                <div style="font-size:0.72rem;font-weight:600;color:#6ee7b7;margin-bottom:8px;text-transform:uppercase;letter-spacing:.05em;">
                  {{ edgePanel.childId }} <span style="font-weight:400;opacity:.7;">(source de données)</span>
                </div>
                <div v-if="!edgePanel.childTables.length"
                     style="font-size:0.78rem;color:var(--text-dim);font-style:italic;padding:4px 0 8px;">
                  Aucune table standard définie.
                </div>
                <div v-for="t in edgePanel.childTables" :key="'c-'+t"
                     style="display:flex;align-items:center;gap:5px;padding:5px 6px;margin-bottom:4px;background:var(--surface);border-radius:5px;font-family:monospace;font-size:0.8rem;">
                  <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{{ t }}</span>
                  <button v-if="!edgePanel.isOneToMany"
                    @click="toggleFluxEntry(t,'child','PULL')"
                    :style="{fontSize:'0.67rem',padding:'2px 7px',borderRadius:'4px',border:'1px solid',cursor:'pointer',transition:'all .15s',
                             background: isActive(t,'child','PULL') ? '#6366f1' : 'transparent',
                             borderColor: isActive(t,'child','PULL') ? '#6366f1' : 'var(--border)',
                             color: isActive(t,'child','PULL') ? '#fff' : 'var(--text-dim)'}">PULL</button>
                  <button @click="toggleFluxEntry(t,'child','COLLECT')"
                    :style="{fontSize:'0.67rem',padding:'2px 7px',borderRadius:'4px',border:'1px solid',cursor:'pointer',transition:'all .15s',
                             background: isActive(t,'child','COLLECT') ? '#d97706' : 'transparent',
                             borderColor: isActive(t,'child','COLLECT') ? '#d97706' : 'var(--border)',
                             color: isActive(t,'child','COLLECT') ? '#fff' : 'var(--text-dim)'}">COL</button>
                </div>
                <!-- Saisie libre child -->
                <div style="display:flex;gap:4px;margin-top:8px;">
                  <input v-model="customChild" placeholder="Autre table / scalaire…"
                    style="flex:1;font-size:0.75rem;padding:4px 7px;background:var(--surface);border:1px solid var(--border);border-radius:4px;color:var(--text);min-width:0;" />
                  <button v-if="!edgePanel.isOneToMany"
                    @click="toggleFluxEntry(customChild,'child','PULL')" :disabled="!customChild"
                    style="font-size:0.67rem;padding:3px 7px;border-radius:4px;border:1px solid var(--border);background:transparent;color:var(--text-dim);cursor:pointer;flex-shrink:0;">PULL</button>
                  <button @click="toggleFluxEntry(customChild,'child','COLLECT')" :disabled="!customChild"
                    style="font-size:0.67rem;padding:3px 7px;border-radius:4px;border:1px solid var(--border);background:transparent;color:var(--text-dim);cursor:pointer;flex-shrink:0;">COL</button>
                </div>
              </div>

              <!-- Section PARENT : config descendante (Dashboard → Posts) -->
              <div style="background:var(--surface2);border-radius:8px;padding:12px;">
                <div style="font-size:0.72rem;font-weight:600;color:#a5b4fc;margin-bottom:8px;text-transform:uppercase;letter-spacing:.05em;">
                  {{ edgePanel.parentId }} <span style="font-weight:400;opacity:.7;">(config → enfants)</span>
                </div>
                <div v-if="!edgePanel.parentTables.length"
                     style="font-size:0.78rem;color:var(--text-dim);font-style:italic;padding:4px 0 8px;">
                  Aucune table standard définie.
                </div>
                <div v-for="t in edgePanel.parentTables" :key="'p-'+t"
                     style="display:flex;align-items:center;gap:5px;padding:5px 6px;margin-bottom:4px;background:var(--surface);border-radius:5px;font-family:monospace;font-size:0.8rem;">
                  <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{{ t }}</span>
                  <button @click="toggleFluxEntry(t,'parent','PUSH')"
                    :style="{fontSize:'0.67rem',padding:'2px 7px',borderRadius:'4px',border:'1px solid',cursor:'pointer',transition:'all .15s',
                             background: isActive(t,'parent','PUSH') ? '#0891b2' : 'transparent',
                             borderColor: isActive(t,'parent','PUSH') ? '#0891b2' : 'var(--border)',
                             color: isActive(t,'parent','PUSH') ? '#fff' : 'var(--text-dim)'}">PUSH</button>
                </div>
                <!-- Saisie libre parent -->
                <div style="display:flex;gap:4px;margin-top:8px;">
                  <input v-model="customParent" placeholder="Autre table / scalaire…"
                    style="flex:1;font-size:0.75rem;padding:4px 7px;background:var(--surface);border:1px solid var(--border);border-radius:4px;color:var(--text);min-width:0;" />
                  <button @click="toggleFluxEntry(customParent,'parent','PUSH')" :disabled="!customParent"
                    style="font-size:0.67rem;padding:3px 7px;border-radius:4px;border:1px solid var(--border);background:transparent;color:var(--text-dim);cursor:pointer;flex-shrink:0;">PUSH</button>
                </div>
              </div>
            </div>

            <!-- Colonne droite : flux déclarés -->
            <div style="background:var(--surface2);border-radius:8px;padding:12px;">
              <div style="font-size:0.75rem;font-weight:600;color:var(--text-dim);margin-bottom:10px;text-transform:uppercase;letter-spacing:.05em;">
                Flux déclarés · {{ edgePanel.rel.flux?.length ?? 0 }}
              </div>
              <div v-if="!edgePanel.rel.flux?.length"
                   style="font-size:0.78rem;color:var(--text-dim);font-style:italic;padding:8px 0;">
                Cliquez un bouton à gauche pour déclarer un flux.
              </div>
              <div v-for="fx in edgePanel.rel.flux" :key="fx.id"
                   style="display:flex;align-items:center;gap:5px;padding:7px 8px;margin-bottom:5px;background:var(--surface);border-radius:6px;font-family:monospace;font-size:0.78rem;">
                <!-- Origine → destination -->
                <span style="font-size:0.67rem;color:var(--text-dim);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px;">
                  {{ (fx.source||'child')==='child' ? edgePanel.childId : edgePanel.parentId }}
                  → {{ (fx.source||'child')==='child' ? edgePanel.parentId : edgePanel.childId }}
                </span>
                <!-- Badge mode -->
                <span :style="{
                  fontSize:'0.64rem',padding:'1px 6px',borderRadius:'3px',fontWeight:'700',flexShrink:0,
                  background: fx.mode==='PULL' ? '#312e81' : fx.mode==='PUSH' ? '#164e63' : '#78350f',
                  color:'#fff'}">{{ fx.mode }}</span>
                <!-- Nom table -->
                <span style="color:var(--accent);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{{ fx.table }}</span>
                <button class="btn btn-danger btn-xs" style="flex-shrink:0;" @click="removeFlux(fx.id)">✕</button>
              </div>
            </div>
          </div>

          <!-- Légende -->
          <div style="margin-top:14px;font-size:0.73rem;color:var(--text-dim);border-top:1px solid var(--border);padding-top:10px;display:flex;gap:16px;flex-wrap:wrap;">
            <span><span style="color:#6366f1;font-weight:600;">PULL</span> — parent lit depuis un enfant</span>
            <span><span style="color:#d97706;font-weight:600;">COL</span> — parent agrège tous les enfants</span>
            <span><span style="color:#0891b2;font-weight:600;">PUSH</span> — parent pousse vers les enfants</span>
            <span style="margin-left:auto;opacity:.6;">Patrons MXL · Pré-remplissage Tissage N2</span>
          </div>
        </div>
      </div>

    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Classes — Master-detail avec 5 onglets
// ═══════════════════════════════════════════════════════════════════════════════

const ViewClasses = {
  components: { TagsInput, FieldEditor, TableStdEditor, MxlHighlight },
  props: ['initialClassId'],
  setup(props) {
    const classes    = ref([]);
    const selected   = ref(null);
    const activeTab  = ref('identite');
    const showCreate = ref(false);
    const mxlPreview = ref('');
    const mxlLoading = ref(false);

    const emptyClass = () => ({
      id: '', label: '', description: '', owner_function: '',
      min_sheets: [], optional_sheets: [], allowed_namespaces: [],
      push_prefix: '', template: '', min_fields: [], std_tables: [],
    });
    const newClass = ref(emptyClass());

    const functions   = ref([]);
    const namespaces  = ref([]);

    const load = async () => {
      [classes.value, functions.value, namespaces.value] = await Promise.all([
        GET('/api/classes').catch(() => []),
        GET('/api/functions').catch(() => []),
        GET('/api/namespaces').catch(() => []),
      ]);
      if (props.initialClassId) {
        const found = classes.value.find(c => c.id === props.initialClassId);
        if (found) selectClass(found);
      }
    };
    onMounted(load);

    watch(() => props.initialClassId, id => {
      if (id) {
        const found = classes.value.find(c => c.id === id);
        if (found) selectClass(found);
      }
    });

    const selectClass = cls => {
      selected.value = JSON.parse(JSON.stringify(cls));
      activeTab.value = 'identite';
      mxlPreview.value = '';
    };

    const save = async () => {
      try {
        if (selected.value.id) {
          await PUT(`/api/classes/${selected.value.id}`, selected.value);
        }
        await load();
        toast('Classe sauvegardée');
      } catch(e) { toastErr(e); }
    };

    const create = async () => {
      try {
        await POST('/api/classes', newClass.value);
        showCreate.value = false;
        newClass.value = emptyClass();
        await load();
        toast('Classe créée');
      } catch(e) { toastErr(e); }
    };

    const deleteClass = async cls => {
      if (!confirm(`Supprimer la Classe "${cls.label}" ?`)) return;
      await DEL(`/api/classes/${cls.id}`).catch(toastErr);
      selected.value = null;
      await load();
    };

    const loadMxl = async () => {
      if (!selected.value) return;
      mxlLoading.value = true;
      try {
        const r = await GET(`/api/mxl/generate/${selected.value.id}`);
        mxlPreview.value = r.mxl;
      } catch(e) { toastErr(e); }
      finally { mxlLoading.value = false; }
    };

    watch(activeTab, t => { if (t === 'mxl' && selected.value) loadMxl(); });

    const toggleNs = (id, checked) => {
      if (!selected.value) return;
      if (checked) {
        if (!selected.value.allowed_namespaces.includes(id))
          selected.value.allowed_namespaces.push(id);
      } else {
        selected.value.allowed_namespaces = selected.value.allowed_namespaces.filter(n => n !== id);
      }
    };

    return {
      classes, functions, namespaces, selected, activeTab, showCreate, newClass,
      mxlPreview, mxlLoading,
      selectClass, save, create, deleteClass, load, toggleNs,
    };
  },
  template: `
    <div class="master-detail" style="height:calc(100vh - 140px)">
      <!-- Liste -->
      <div class="master-list">
        <div style="padding:8px 12px;display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
          <span style="font-size:0.72rem;color:var(--text-dim);text-transform:uppercase;letter-spacing:.08em">Classes</span>
          <button class="btn btn-primary btn-xs" @click="showCreate=true">+</button>
        </div>
        <div v-for="cls in classes" :key="cls.id"
             class="master-item" :class="{selected: selected && selected.id===cls.id}"
             @click="selectClass(cls)">
          <div class="master-item-label">{{ cls.label || cls.id }}</div>
          <div class="master-item-sub">{{ cls.id }}</div>
        </div>
        <div v-if="!classes.length" class="empty" style="padding:24px;font-size:0.78rem">Aucune Classe</div>
      </div>

      <!-- Détail -->
      <div class="detail-panel" v-if="selected">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
          <span style="font-weight:700;font-size:1rem">{{ selected.label || selected.id }}</span>
          <div style="display:flex;gap:8px">
            <button class="btn btn-danger btn-sm" @click="deleteClass(selected)">Supprimer</button>
            <button class="btn btn-primary btn-sm" @click="save">💾 Sauvegarder</button>
          </div>
        </div>

        <div class="tabs">
          <button v-for="t in [{k:'identite',l:'Identité'},{k:'structure',l:'Structure'},{k:'champs',l:'Champs'},{k:'tables',l:'Tables std'},{k:'mxl',l:'Aperçu MXL'}]"
                  :key="t.k" class="tab-btn" :class="{active:activeTab===t.k}" @click="activeTab=t.k">{{ t.l }}</button>
        </div>

        <!-- Onglet Identité -->
        <div v-if="activeTab==='identite'">
          <div class="form-row">
            <div class="form-group">
              <label>ID (identifiant technique)</label>
              <input v-model="selected.id" disabled style="opacity:.6;cursor:not-allowed" />
            </div>
            <div class="form-group">
              <label>Libellé</label>
              <input v-model="selected.label" placeholder="Unité d'Œuvre" />
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="selected.description" placeholder="Rôle de cette Classe dans l'écosystème…" style="min-height:60px"></textarea>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Fonction owner (rôle requis)</label>
              <select v-model="selected.owner_function">
                <option value="">— aucune —</option>
                <option v-for="fn in functions" :key="fn.id" :value="fn.id">
                  {{ fn.label || fn.id }}
                  <template v-if="fn.side"> ({{ fn.side }})</template>
                </option>
              </select>
            </div>
            <div class="form-group">
              <label>Push prefix (namespace)</label>
              <input v-model="selected.push_prefix" placeholder="uo.{id}." />
            </div>
          </div>
        </div>

        <!-- Onglet Structure -->
        <div v-if="activeTab==='structure'">
          <div class="form-group">
            <label>Feuilles obligatoires (min_sheets)</label>
            <tags-input v-model="selected.min_sheets" placeholder="Ajouter feuille…" />
          </div>
          <div class="form-group">
            <label>Feuilles optionnelles</label>
            <tags-input v-model="selected.optional_sheets" placeholder="Dashboard, _Log…" />
          </div>
          <div class="form-group">
            <label>Namespaces autorisés</label>
            <div v-if="namespaces.length" style="display:flex;flex-wrap:wrap;gap:10px;margin-top:6px">
              <label v-for="ns in namespaces" :key="ns.id"
                     style="display:flex;align-items:center;gap:6px;font-size:0.83rem;cursor:pointer;padding:4px 8px;border:1px solid var(--border);border-radius:4px"
                     :style="selected.allowed_namespaces.includes(ns.id) ? 'border-color:var(--accent);background:rgba(99,102,241,.08)' : ''">
                <input type="checkbox"
                       :checked="selected.allowed_namespaces.includes(ns.id)"
                       @change="toggleNs(ns.id, $event.target.checked)"
                       style="width:auto;accent-color:var(--accent)" />
                <span :title="ns.description">{{ ns.label || ns.id }}</span>
                <span v-if="ns.prefix" style="font-family:monospace;color:var(--accent2);font-size:0.73rem">{{ ns.prefix }}</span>
              </label>
            </div>
            <div v-else style="font-size:0.78rem;color:var(--text-dim);margin-top:4px">
              Aucun namespace défini — créez-en dans la vue <em>Namespaces</em>.
            </div>
          </div>
          <div class="form-group">
            <label>Template MXL (chemin)</label>
            <input v-model="selected.template" placeholder="config/templates/…" />
          </div>
        </div>

        <!-- Onglet Champs -->
        <div v-if="activeTab==='champs'">
          <p style="font-size:0.78rem;color:var(--text-dim);margin-bottom:12px">
            Champs scalaires (plages nommées Excel) — déclarés en <span style="color:var(--accent2);font-family:monospace">DEF … = GET_CELL()</span>
          </p>
          <field-editor v-model="selected.min_fields" />
        </div>

        <!-- Onglet Tables std -->
        <div v-if="activeTab==='tables'">
          <p style="font-size:0.78rem;color:var(--text-dim);margin-bottom:12px">
            Tables Excel standard pour cette Classe — déclarées en <span style="color:var(--accent2);font-family:monospace">DEF … = GET_TABLE()</span>
          </p>
          <table-std-editor v-model="selected.std_tables" :writeFunctions="functions" />
        </div>

        <!-- Onglet Aperçu MXL -->
        <div v-if="activeTab==='mxl'">
          <div style="display:flex;justify-content:flex-end;margin-bottom:10px">
            <button class="btn btn-ghost btn-sm" @click="loadMxl">⟳ Régénérer</button>
          </div>
          <div v-if="mxlLoading" class="loading">Génération…</div>
          <mxl-highlight v-else :code="mxlPreview" />
        </div>
      </div>
      <div v-else class="detail-panel empty">Sélectionnez une Classe dans la liste</div>

      <!-- Modal création -->
      <div v-if="showCreate" class="modal-overlay" @click.self="showCreate=false">
        <div class="modal">
          <div class="modal-title">Nouvelle Classe</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID (identifiant technique unique)</label>
              <input v-model="newClass.id" placeholder="uo_instance" autofocus />
            </div>
            <div class="form-group">
              <label>Libellé</label>
              <input v-model="newClass.label" placeholder="Unité d'Œuvre" />
            </div>
          </div>
          <div class="form-group">
            <label>Fonction owner</label>
            <select v-model="newClass.owner_function">
              <option value="">— aucune —</option>
              <option v-for="fn in functions" :key="fn.id" :value="fn.id">
                {{ fn.label || fn.id }}
                <template v-if="fn.side"> ({{ fn.side }})</template>
              </option>
            </select>
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showCreate=false">Annuler</button>
            <button class="btn btn-primary" @click="create" :disabled="!newClass.id">Créer</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Relations P/F
// ═══════════════════════════════════════════════════════════════════════════════

const ViewRelations = {
  setup() {
    const relations = ref([]);
    const classes   = ref([]);
    const showModal = ref(false);
    const editing   = ref(null);

    const empty = () => ({ id:'', parent_class:'', child_class:'', qualifier:'TYPICAL', cardinality:'1..N', description:'' });

    const load = async () => {
      [relations.value, classes.value] = await Promise.all([
        GET('/api/relations').catch(() => []),
        GET('/api/classes').catch(() => []),
      ]);
    };
    onMounted(load);

    const openCreate = () => { editing.value = empty(); showModal.value = true; };
    const openEdit   = r  => { editing.value = {...r};   showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value.id) {
          await PUT(`/api/relations/${editing.value.id}`, editing.value);
        } else {
          await POST('/api/relations', editing.value);
        }
        showModal.value = false;
        await load();
        toast('Relation sauvegardée');
      } catch(e) { toastErr(e); }
    };

    const del = async r => {
      if (!confirm('Supprimer cette relation ?')) return;
      await DEL(`/api/relations/${r.id}`).catch(toastErr);
      await load();
    };

    const toggleQualifier = async r => {
      const updated = {...r, qualifier: r.qualifier === 'PRESCRIBED' ? 'TYPICAL' : 'PRESCRIBED'};
      await PUT(`/api/relations/${r.id}`, updated).catch(toastErr);
      await load();
    };

    const classLabel = id => (classes.value.find(c => c.id === id) || {}).label || id;

    return { relations, classes, showModal, editing, classLabel, openCreate, openEdit, save, del, toggleQualifier };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:16px">
        <span class="card-title">Relations Père / Fils</span>
        <button class="btn btn-primary btn-sm" @click="openCreate">+ Relation</button>
      </div>
      <div class="card" style="padding:0">
        <table>
          <thead><tr><th>Père</th><th>Fils</th><th>Qualifier</th><th>Cardinalité</th><th>Description</th><th></th></tr></thead>
          <tbody>
            <tr v-for="r in relations" :key="r.id">
              <td><span class="badge badge-indigo">{{ classLabel(r.parent_class) }}</span></td>
              <td><span class="badge badge-gray">{{ classLabel(r.child_class) }}</span></td>
              <td>
                <button class="btn btn-xs" @click="toggleQualifier(r)"
                  :style="r.qualifier==='PRESCRIBED' ? 'color:#fca5a5;border-color:#ef4444;background:transparent' : 'color:#94a3b8;border-color:#475569;background:transparent'">
                  {{ r.qualifier }}
                </button>
              </td>
              <td style="color:var(--text-dim);font-size:0.8rem">{{ r.cardinality }}</td>
              <td style="color:var(--text-dim);font-size:0.78rem">{{ r.description }}</td>
              <td>
                <button class="btn btn-ghost btn-xs" @click="openEdit(r)" style="margin-right:6px">✏</button>
                <button class="btn btn-danger btn-xs" @click="del(r)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
        <div v-if="!relations.length" class="empty">Aucune relation définie</div>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing && editing.id ? 'Éditer' : 'Nouvelle' }} Relation P/F</div>
          <div class="form-row">
            <div class="form-group">
              <label>Classe Père</label>
              <select v-model="editing.parent_class">
                <option value="">— choisir —</option>
                <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label || c.id }}</option>
              </select>
            </div>
            <div class="form-group">
              <label>Classe Fils</label>
              <select v-model="editing.child_class">
                <option value="">— choisir —</option>
                <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label || c.id }}</option>
              </select>
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Qualifier</label>
              <select v-model="editing.qualifier">
                <option value="TYPICAL">TYPICAL (tirets)</option>
                <option value="PRESCRIBED">PRESCRIBED (obligatoire)</option>
              </select>
            </div>
            <div class="form-group">
              <label>Cardinalité</label>
              <input v-model="editing.cardinality" placeholder="1..N" />
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <input v-model="editing.description" placeholder="Description optionnelle…" />
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">Sauvegarder</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Fonctions Acteurs
// ═══════════════════════════════════════════════════════════════════════════════

const ViewFunctions = {
  setup() {
    const functions = ref([]);
    const showModal = ref(false);
    const editing   = ref(null);

    const empty = () => ({ id:'', label:'', description:'', side:'interne' });

    const load = async () => { functions.value = await GET('/api/functions').catch(() => []); };
    onMounted(load);

    const openCreate = () => { editing.value = empty(); showModal.value = true; };
    const openEdit   = f  => { editing.value = {...f};  showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value.id && functions.value.find(f => f.id === editing.value.id)) {
          await PUT(`/api/functions/${editing.value.id}`, editing.value);
        } else {
          await POST('/api/functions', editing.value);
        }
        showModal.value = false;
        await load();
        toast('Fonction sauvegardée');
      } catch(e) { toastErr(e); }
    };

    const del = async f => {
      if (!confirm(`Supprimer la Fonction "${f.label}" ?`)) return;
      await DEL(`/api/functions/${f.id}`).catch(toastErr);
      await load();
    };

    return { functions, showModal, editing, openCreate, openEdit, save, del };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:16px">
        <span class="card-title">Fonctions Acteurs</span>
        <button class="btn btn-primary btn-sm" @click="openCreate">+ Fonction</button>
      </div>
      <div class="card" style="padding:0">
        <table>
          <thead><tr><th>ID</th><th>Libellé</th><th>Side</th><th>Description</th><th></th></tr></thead>
          <tbody>
            <tr v-for="f in functions" :key="f.id">
              <td style="font-family:monospace;font-size:0.78rem;color:var(--text-dim)">{{ f.id }}</td>
              <td style="font-weight:500">{{ f.label }}</td>
              <td>
                <span class="badge" :class="f.side==='interne' ? 'badge-interne' : 'badge-externe'">{{ f.side }}</span>
              </td>
              <td style="color:var(--text-dim);font-size:0.78rem">{{ f.description }}</td>
              <td>
                <button class="btn btn-ghost btn-xs" @click="openEdit(f)" style="margin-right:6px">✏</button>
                <button class="btn btn-danger btn-xs" @click="del(f)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
        <div v-if="!functions.length" class="empty">Aucune Fonction définie</div>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal">
          <div class="modal-title">{{ editing && editing.id ? 'Éditer' : 'Nouvelle' }} Fonction</div>
          <div class="form-row">
            <div class="form-group">
              <label>ID</label>
              <input v-model="editing.id" placeholder="ingenieur_sys" :disabled="!!editing._saved" />
            </div>
            <div class="form-group">
              <label>Libellé</label>
              <input v-model="editing.label" placeholder="Ingénieur système" />
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label>Side</label>
              <select v-model="editing.side">
                <option value="interne">interne</option>
                <option value="externe">externe</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="editing.description" style="min-height:60px" placeholder="Rôle dans l'écosystème…"></textarea>
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">Sauvegarder</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Templates
// ═══════════════════════════════════════════════════════════════════════════════

const ViewTemplates = {
  components: { TableStdEditor },
  setup() {
    const templates  = ref([]);
    const classes    = ref([]);
    const filterCls  = ref('');
    const showModal  = ref(false);
    const editing    = ref(null);

    const empty = () => ({ id:'', label:'', class_id:'', description:'', extra_sheets:[], field_defaults:{}, std_tables:[], mxl_defaults:{}, source_file:'' });

    const functions = ref([]);

    const load = async () => {
      [templates.value, classes.value, functions.value] = await Promise.all([
        GET('/api/templates').catch(() => []),
        GET('/api/classes').catch(() => []),
        GET('/api/functions').catch(() => []),
      ]);
    };
    onMounted(load);

    const filtered = computed(() => filterCls.value ? templates.value.filter(t => t.class_id === filterCls.value) : templates.value);
    const classLabel = id => (classes.value.find(c => c.id === id) || {}).label || id;

    const openCreate = () => { editing.value = empty(); showModal.value = true; };
    const openEdit   = t  => { editing.value = JSON.parse(JSON.stringify(t)); showModal.value = true; };

    const save = async () => {
      try {
        if (editing.value.id && templates.value.find(t => t.id === editing.value.id)) {
          await PUT(`/api/templates/${editing.value.id}`, editing.value);
        } else {
          await POST('/api/templates', editing.value);
        }
        showModal.value = false;
        await load();
        toast('Template sauvegardé');
      } catch(e) { toastErr(e); }
    };

    const del = async t => {
      if (!confirm(`Supprimer le Template "${t.label}" ?`)) return;
      await DEL(`/api/templates/${t.id}`).catch(toastErr);
      await load();
    };

    return { templates, classes, functions, filterCls, filtered, showModal, editing, classLabel, openCreate, openEdit, save, del };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:16px">
        <div style="display:flex;align-items:center;gap:12px">
          <span class="card-title">Templates</span>
          <select v-model="filterCls" style="width:180px;background:var(--surface2);border-color:var(--border)">
            <option value="">Toutes les Classes</option>
            <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label || c.id }}</option>
          </select>
        </div>
        <button class="btn btn-primary btn-sm" @click="openCreate">+ Template</button>
      </div>
      <div class="card" style="padding:0">
        <table>
          <thead><tr><th>ID</th><th>Libellé</th><th>Classe</th><th>Feuilles extra</th><th></th></tr></thead>
          <tbody>
            <tr v-for="t in filtered" :key="t.id">
              <td style="font-family:monospace;font-size:0.78rem;color:var(--text-dim)">{{ t.id }}</td>
              <td style="font-weight:500">{{ t.label }}</td>
              <td><span class="badge badge-indigo">{{ classLabel(t.class_id) }}</span></td>
              <td style="color:var(--text-dim);font-size:0.78rem">{{ (t.extra_sheets||[]).join(', ') || '—' }}</td>
              <td>
                <button class="btn btn-ghost btn-xs" @click="openEdit(t)" style="margin-right:6px">✏</button>
                <button class="btn btn-danger btn-xs" @click="del(t)">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
        <div v-if="!filtered.length" class="empty">Aucun Template</div>
      </div>

      <div v-if="showModal" class="modal-overlay" @click.self="showModal=false">
        <div class="modal modal-lg">
          <div class="modal-title">{{ editing && editing.id ? 'Éditer' : 'Nouveau' }} Template</div>
          <div class="form-row">
            <div class="form-group">
              <label>Libellé</label>
              <input v-model="editing.label" placeholder="Template UO Climatisation" autofocus />
            </div>
            <div class="form-group">
              <label>Classe associée</label>
              <select v-model="editing.class_id">
                <option value="">— choisir —</option>
                <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label || c.id }}</option>
              </select>
            </div>
          </div>
          <div class="form-group">
            <label>Description</label>
            <textarea v-model="editing.description" style="min-height:50px" placeholder="Cas d'usage de ce template…"></textarea>
          </div>
          <div class="form-group">
            <label>Tables supplémentaires</label>
            <table-std-editor v-model="editing.std_tables" :writeFunctions="functions" />
          </div>
          <div class="form-actions">
            <button class="btn btn-ghost" @click="showModal=false">Annuler</button>
            <button class="btn btn-primary" @click="save">Sauvegarder</button>
          </div>
        </div>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Générateur MXL
// ═══════════════════════════════════════════════════════════════════════════════

const ViewMxlPreview = {
  components: { MxlHighlight },
  setup() {
    const classes    = ref([]);
    const selectedId = ref('');
    const mxl        = ref('');
    const loading    = ref(false);

    onMounted(async () => {
      classes.value = await GET('/api/classes').catch(() => []);
    });

    const generate = async () => {
      if (!selectedId.value) return;
      loading.value = true;
      try {
        const r = await GET(`/api/mxl/generate/${selectedId.value}`);
        mxl.value = r.mxl;
      } catch(e) { toastErr(e); mxl.value = ''; }
      finally { loading.value = false; }
    };

    const copy = () => {
      navigator.clipboard.writeText(mxl.value).then(() => toast('MXL copié dans le presse-papiers'));
    };

    return { classes, selectedId, mxl, loading, generate, copy };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:20px">
        <span class="card-title">Générateur de template MXL</span>
      </div>
      <div style="display:flex;gap:12px;align-items:center;margin-bottom:20px">
        <select v-model="selectedId" style="flex:1;max-width:320px">
          <option value="">— Choisir une Classe —</option>
          <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label || c.id }}</option>
        </select>
        <button class="btn btn-primary" @click="generate" :disabled="!selectedId||loading">
          {{ loading ? '…' : '⚡ Générer' }}
        </button>
        <button v-if="mxl" class="btn btn-ghost" @click="copy">📋 Copier</button>
      </div>
      <div v-if="loading" class="loading">Génération du template MXL…</div>
      <mxl-highlight v-else-if="mxl" :code="mxl" />
      <div v-else class="empty" style="padding:40px">
        Sélectionnez une Classe pour générer son template MXL
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Namespaces (CRUD)
// ═══════════════════════════════════════════════════════════════════════════════

const ViewNamespaces = {
  setup() {
    const namespaces = ref([]);
    const showForm   = ref(false);
    const editing    = ref(null);   // id du namespace en édition, ou null
    const form       = ref({ id: '', label: '', prefix: '', description: '' });

    const emptyForm  = () => ({ id: '', label: '', prefix: '', description: '' });

    const load = async () => {
      namespaces.value = await GET('/api/namespaces').catch(() => []);
    };
    onMounted(load);

    const startCreate = () => {
      form.value = emptyForm();
      editing.value = null;
      showForm.value = true;
    };

    const startEdit = ns => {
      form.value = { ...ns };
      editing.value = ns.id;
      showForm.value = true;
    };

    const save = async () => {
      try {
        if (editing.value) {
          await PUT(`/api/namespaces/${editing.value}`, form.value);
          toast('Namespace mis à jour');
        } else {
          await POST('/api/namespaces', form.value);
          toast('Namespace créé');
        }
        showForm.value = false;
        editing.value = null;
        await load();
      } catch(e) { toastErr(e); }
    };

    const del = async ns => {
      if (!confirm(`Supprimer le namespace "${ns.label || ns.id}" ?`)) return;
      await DEL(`/api/namespaces/${ns.id}`).catch(toastErr);
      await load();
    };

    const cancel = () => { showForm.value = false; editing.value = null; };

    return { namespaces, form, editing, showForm, startCreate, startEdit, save, del, cancel };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:16px">
        <span class="card-title">Namespaces de l'écosystème</span>
        <button class="btn btn-primary btn-sm" @click="startCreate">+ Nouveau</button>
      </div>
      <p style="font-size:0.78rem;color:var(--text-dim);margin-bottom:20px">
        Un namespace déclare un préfixe d'identifiant au niveau de l'écosystème.
        Les Classes peuvent ensuite se voir associer un ou plusieurs namespaces autorisés.
      </p>

      <!-- Formulaire création / édition -->
      <div v-if="showForm" class="card" style="margin-bottom:20px">
        <div class="card-title" style="margin-bottom:12px">{{ editing ? 'Modifier' : 'Nouveau' }} namespace</div>
        <div class="form-row">
          <div class="form-group">
            <label>ID technique</label>
            <input v-model="form.id" :disabled="!!editing"
                   :style="editing ? 'opacity:.6;cursor:not-allowed' : ''"
                   placeholder="uo" />
          </div>
          <div class="form-group">
            <label>Libellé</label>
            <input v-model="form.label" placeholder="Unités Opérationnelles" />
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>Préfixe</label>
            <input v-model="form.prefix" placeholder="uo." style="font-family:monospace" />
          </div>
          <div class="form-group">
            <label>Description</label>
            <input v-model="form.description" placeholder="Description optionnelle…" />
          </div>
        </div>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:8px">
          <button class="btn btn-ghost btn-sm" @click="cancel">Annuler</button>
          <button class="btn btn-primary btn-sm" @click="save">💾 Sauvegarder</button>
        </div>
      </div>

      <!-- Liste -->
      <div v-if="namespaces.length">
        <table>
          <thead>
            <tr>
              <th>ID</th><th>Libellé</th><th>Préfixe</th><th>Description</th><th></th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="ns in namespaces" :key="ns.id">
              <td style="font-family:monospace;font-size:0.82rem">{{ ns.id }}</td>
              <td>{{ ns.label }}</td>
              <td style="font-family:monospace;color:var(--accent2)">{{ ns.prefix }}</td>
              <td style="font-size:0.78rem;color:var(--text-dim)">{{ ns.description }}</td>
              <td style="text-align:right;white-space:nowrap">
                <button class="btn btn-ghost btn-xs" @click="startEdit(ns)">✏</button>
                <button class="btn btn-danger btn-xs" @click="del(ns)" style="margin-left:4px">✕</button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div v-else-if="!showForm" style="color:var(--text-dim);font-size:0.78rem;padding:32px;text-align:center">
        Aucun namespace défini dans cet écosystème.
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Import Excel → Classe (Wizard 3 étapes)
// ═══════════════════════════════════════════════════════════════════════════════

const ViewImportExcel = {
  setup() {
    const step            = ref(1);
    const scan            = ref(null);
    const classes         = ref([]);
    const preview         = ref(null);
    const targetId        = ref('');
    const createNew       = ref(false);
    const newLabel        = ref('');
    const uploading       = ref(false);
    const applying        = ref(false);
    const done            = ref(null);
    const sheetSelections = ref({});   // sheet → 'obligatoire' | 'optionnelle' | 'ignorer'

    onMounted(async () => {
      classes.value = await GET('/api/classes').catch(() => []);
    });

    // Pré-sélection automatique des feuilles après le scan
    const computeSheetSelections = () => {
      if (!scan.value) return;
      const usedSheets = new Set(scan.value.tables.map(t => t.sheet));
      const sels = {};
      scan.value.sheets.forEach(s => {
        if (s.startsWith('_'))        sels[s] = 'ignorer';
        else if (usedSheets.has(s))   sels[s] = 'obligatoire';
        else                          sels[s] = 'optionnelle';
      });
      sheetSelections.value = sels;
    };

    const onFileDrop = async e => {
      const file = (e.dataTransfer || e.target).files[0];
      if (!file) return;
      uploading.value = true;
      try {
        const fd = new FormData();
        fd.append('file', file);
        const r = await fetch('/api/excel/scan-upload', { method: 'POST', body: fd });
        if (!r.ok) throw new Error((await r.json()).detail);
        scan.value = await r.json();
        computeSheetSelections();
        step.value = 2;
        await buildPreview();
      } catch(e) { toastErr(e); }
      finally { uploading.value = false; }
    };

    const buildPreview = async () => {
      if (!scan.value) return;
      preview.value = await POST('/api/excel/preview-class', { scan: scan.value, class_id: targetId.value }).catch(e => { toastErr(e); return null; });
    };

    watch(targetId, buildPreview);

    const apply = async () => {
      applying.value = true;
      try {
        const classId = targetId.value || preview.value?.class_id || 'nouvelle_classe';
        const sel = sheetSelections.value;
        const r = await POST('/api/excel/apply-class', {
          class_id: classId,
          label: newLabel.value || classId,
          field_mappings: preview.value?.suggested_fields || [],
          table_mappings: preview.value?.suggested_tables || [],
          min_sheets:      Object.entries(sel).filter(([,v]) => v === 'obligatoire').map(([k]) => k),
          optional_sheets: Object.entries(sel).filter(([,v]) => v === 'optionnelle').map(([k]) => k),
          create_if_missing: true,
        });
        done.value = r;
        step.value = 3;
        classes.value = await GET('/api/classes').catch(() => []);
      } catch(e) { toastErr(e); }
      finally { applying.value = false; }
    };

    const reset = () => {
      step.value=1; scan.value=null; preview.value=null; targetId.value='';
      done.value=null; createNew.value=false; newLabel.value=''; sheetSelections.value={};
    };

    const toggleField = (f, val) => { f.include = val; };
    const toggleTable = (t, val) => { t.include = val; };

    return { step, scan, classes, preview, targetId, createNew, newLabel, uploading, applying, done,
             sheetSelections, onFileDrop, apply, reset, toggleField, toggleTable };
  },
  template: `
    <div>
      <div class="card-header" style="margin-bottom:20px">
        <span class="card-title">Import Excel → Classe</span>
        <button v-if="step>1" class="btn btn-ghost btn-sm" @click="reset">⟵ Recommencer</button>
      </div>

      <!-- Steps indicator -->
      <div class="wizard-steps" style="max-width:500px;margin-bottom:24px">
        <div class="wizard-step" :class="{active:step===1,done:step>1}">
          <div><div class="step-num">1</div></div>Upload Excel
        </div>
        <div class="wizard-step" :class="{active:step===2,done:step>2}">
          <div><div class="step-num">2</div></div>Mapping
        </div>
        <div class="wizard-step" :class="{active:step===3}">
          <div><div class="step-num">3</div></div>Appliquer
        </div>
      </div>

      <!-- Step 1: Upload -->
      <div v-if="step===1">
        <div class="upload-zone" @dragover.prevent @drop.prevent="onFileDrop" :class="{drag:uploading}">
          <div style="font-size:2rem;margin-bottom:12px">📄</div>
          <div style="font-weight:600;margin-bottom:8px">Glissez un fichier Excel ici</div>
          <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px">.xlsx, .xlsm</div>
          <label class="btn btn-ghost" style="cursor:pointer">
            Parcourir
            <input type="file" accept=".xlsx,.xlsm" style="display:none" @change="onFileDrop" />
          </label>
          <div v-if="uploading" style="margin-top:12px;color:var(--accent2)">Analyse en cours…</div>
        </div>
      </div>

      <!-- Step 2: Mapping -->
      <div v-if="step===2 && scan">
        <div class="card" style="margin-bottom:16px">
          <div class="card-header" style="margin-bottom:10px">
            <span class="card-title">Fichier : {{ scan.filename }}</span>
            <span class="badge badge-gray">{{ scan.sheets.length }} feuille(s)</span>
          </div>
          <div style="font-size:0.78rem;color:var(--text-dim)">
            {{ scan.tables.length }} table(s) Excel · {{ scan.named_ranges.filter(r=>r.kind==='scalar').length }} plage(s) nommée(s) scalaire(s)
          </div>
        </div>

        <div class="form-row" style="margin-bottom:20px">
          <div class="form-group">
            <label>Classe cible</label>
            <select v-model="targetId">
              <option value="">— Créer une nouvelle Classe —</option>
              <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.label || c.id }}</option>
            </select>
          </div>
          <div v-if="!targetId" class="form-group">
            <label>Nom de la nouvelle Classe</label>
            <input v-model="newLabel" placeholder="Identifiant (ex: cockpit_new)" />
          </div>
        </div>

        <!-- Feuilles -->
        <div class="card" style="margin-bottom:16px">
          <div class="card-title" style="margin-bottom:12px">Feuilles détectées</div>
          <table>
            <thead><tr><th>Feuille</th><th>Rôle dans la Classe</th><th>Tables liées</th></tr></thead>
            <tbody>
              <tr v-for="sheet in scan.sheets" :key="sheet">
                <td style="font-family:monospace;font-size:0.82rem">{{ sheet }}</td>
                <td>
                  <select :value="sheetSelections[sheet]"
                          @change="sheetSelections[sheet]=$event.target.value"
                          style="width:150px">
                    <option value="obligatoire">Obligatoire</option>
                    <option value="optionnelle">Optionnelle</option>
                    <option value="ignorer">Ignorer</option>
                  </select>
                </td>
                <td style="font-size:0.75rem;color:var(--text-dim)">
                  {{ (preview?.suggested_tables||[]).filter(t=>t.sheet===sheet&&t.include).map(t=>t.excel_name).join(', ') || '—' }}
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <div v-if="preview">
          <!-- Champs scalaires -->
          <div class="card" style="margin-bottom:16px">
            <div class="card-title" style="margin-bottom:12px">Champs scalaires détectés</div>
            <table>
              <thead><tr><th>Plage nommée</th><th>Nom champ</th><th>Type</th><th>Inclure</th></tr></thead>
              <tbody>
                <tr v-for="f in preview.suggested_fields" :key="f.named_range">
                  <td style="font-family:monospace;font-size:0.78rem">{{ f.named_range }}</td>
                  <td><input :value="f.field_name" @input="f.field_name=$event.target.value" style="width:160px" /></td>
                  <td>
                    <select :value="f.field_type" @change="f.field_type=$event.target.value" style="width:90px">
                      <option v-for="t in ['string','int','float','date','bool']" :key="t" :value="t">{{ t }}</option>
                    </select>
                  </td>
                  <td style="text-align:center">
                    <input type="checkbox" :checked="f.include" @change="f.include=$event.target.checked" style="width:auto;accent-color:var(--accent)" />
                  </td>
                </tr>
                <tr v-if="!preview.suggested_fields.length">
                  <td colspan="4" style="color:var(--text-dim);font-size:0.78rem">Aucune plage nommée scalaire trouvée</td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Tables -->
          <div class="card" style="margin-bottom:20px">
            <div class="card-title" style="margin-bottom:12px">Tables Excel détectées</div>
            <table>
              <thead><tr><th>Table Excel</th><th>Nom schéma</th><th>Feuille</th><th>Colonnes</th><th>Inclure</th></tr></thead>
              <tbody>
                <tr v-for="t in preview.suggested_tables" :key="t.excel_name">
                  <td style="font-family:monospace;font-size:0.78rem">{{ t.excel_name }}</td>
                  <td><input :value="t.schema_name" @input="t.schema_name=$event.target.value" style="width:150px" /></td>
                  <td style="color:var(--text-dim);font-size:0.78rem">{{ t.sheet }}</td>
                  <td style="color:var(--text-dim);font-size:0.78rem">{{ t.columns.length }}</td>
                  <td style="text-align:center">
                    <input type="checkbox" :checked="t.include" @change="t.include=$event.target.checked" style="width:auto;accent-color:var(--accent)" />
                  </td>
                </tr>
                <tr v-if="!preview.suggested_tables.length">
                  <td colspan="5" style="color:var(--text-dim);font-size:0.78rem">Aucune table Excel trouvée</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <div style="display:flex;justify-content:flex-end">
          <button class="btn btn-primary" @click="apply" :disabled="applying">
            {{ applying ? 'Application…' : '✓ Appliquer à la Classe' }}
          </button>
        </div>
      </div>

      <!-- Step 3: Résultat -->
      <div v-if="step===3 && done" style="text-align:center;padding:40px">
        <div style="font-size:2.5rem;margin-bottom:16px">✓</div>
        <div style="font-size:1.1rem;font-weight:600;color:var(--ok);margin-bottom:12px">Classe mise à jour</div>
        <div style="font-size:0.85rem;color:var(--text-dim);margin-bottom:24px">
          <strong style="color:var(--text)">{{ done.class_id }}</strong> —
          {{ done.fields_added }} champ(s) ajouté(s) · {{ done.tables_added }} table(s) ajoutée(s)
        </div>
        <button class="btn btn-primary" @click="reset">Importer un autre fichier</button>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Workspace global (partagé N1 + N2)
// ═══════════════════════════════════════════════════════════════════════════════

const ViewWorkspace = {
  setup() {
    const form    = ref({ gabarits_dir: '' });
    const loading = ref(true);
    const saved   = ref(false);
    const err     = ref('');

    onMounted(async () => {
      try {
        const d = await GET('/api/workspace');
        form.value = { gabarits_dir: d.gabarits_dir || '' };
      } catch(e) { toastErr(e); }
      loading.value = false;
    });

    const save = async () => {
      err.value = '';
      try {
        await PUT('/api/workspace', form.value);
        saved.value = true;
        setTimeout(() => saved.value = false, 2000);
        toast('Répertoire des Gabarits enregistré');
      } catch(e) { err.value = e.message; toastErr(e); }
    };

    const reset = async () => {
      if (!confirm('Réinitialiser le répertoire des Gabarits ?')) return;
      await DEL('/api/workspace');
      form.value = { gabarits_dir: '' };
      toast('Répertoire des Gabarits réinitialisé');
    };

    return { form, loading, saved, err, save, reset };
  },
  template: `
    <div v-if="loading" class="loading">Chargement…</div>
    <div v-else style="max-width:640px;">

      <div style="margin-bottom:20px;padding:14px 16px;background:var(--surface2);border-radius:8px;border-left:3px solid #a5b4fc;font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        N1 gère le <strong style="color:var(--text)">Répertoire des Gabarits</strong> —
        les schémas réutilisables copiés par N2 pour créer des Affaires.
        Stocké dans <code style="color:#a5b4fc;">.exosync_workspace.json</code> (partagé avec N2).
      </div>

      <div class="form-group">
        <label class="form-label">Répertoire des Gabarits</label>
        <input v-model="form.gabarits_dir" class="form-control"
          placeholder="ex: C:\\\\ExoSync\\\\Gabarits  ou  /home/user/gabarits" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Chemin absolu. Chaque sous-dossier contenant <code>file_types.yaml</code> est détecté comme un gabarit.
          Les nouveaux gabarits créés ici seront placés automatiquement sous ce dossier.
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
        <div style="font-weight:600;color:var(--text);margin-bottom:8px;">Structure attendue</div>
        <pre style="font-family:monospace;font-size:0.75rem;line-height:1.7;margin:0;color:var(--text-dim);">{{ form.gabarits_dir || 'gabarits_dir' }}/
  SysEng/                 ← gabarit créé depuis N1
    file_types.yaml       ← requis
    schema_relations.json
    functions.json
  SysEng-v2/
    file_types.yaml</pre>
      </div>
    </div>
  `
};


// ═══════════════════════════════════════════════════════════════════════════════
// APP ROOT
// ═══════════════════════════════════════════════════════════════════════════════

const App = {
  components: {
    ToastLayer,
    ViewEcosystemManager, ViewWorkspace, ViewBlueprint, ViewClasses,
    ViewRelations, ViewFunctions, ViewTemplates,
    ViewNamespaces, ViewMxlPreview, ViewImportExcel,
  },
  setup() {
    const view         = ref('ecosystem');
    const ecoName      = ref('');
    const editClassId  = ref(null);

    const loadEco = async () => {
      try {
        const eco = await GET('/api/ecosystem/current');
        ecoName.value = eco.name;
      } catch { ecoName.value = 'non chargé'; }
    };
    onMounted(loadEco);

    const onEcoOpen = name => {
      ecoName.value = name;
      view.value = 'blueprint';
    };

    const navTo = v => {
      view.value = v;
      editClassId.value = null;
    };

    const onEditClass = id => {
      editClassId.value = id;
      view.value = 'classes';
    };

    const NAV = [
      { group: 'Espace de travail', items: [
        { key: 'ecosystem',  icon: '⬡', label: 'Écosystèmes' },
        { key: 'workspace',  icon: '⌂', label: 'Workspace global' },
        { key: 'blueprint',  icon: '◎', label: 'Blueprint' },
      ]},
      { group: 'Schéma N1', items: [
        { key: 'classes',    icon: '◻', label: 'Classes' },
        { key: 'relations',  icon: '↔', label: 'Relations P/F' },
        { key: 'functions',  icon: '◈', label: 'Fonctions' },
        { key: 'namespaces', icon: '⬡', label: 'Namespaces' },
        { key: 'templates',  icon: '⊞', label: 'Templates' },
      ]},
      { group: 'Outils', items: [
        { key: 'mxl',    icon: '⌨', label: 'Générateur MXL' },
        { key: 'import', icon: '↑', label: 'Import Excel' },
      ]},
    ];

    const VIEW_TITLES = {
      ecosystem:  'Gestionnaire d\'Écosystèmes',
      workspace:  'Workspace Global — Gabarits & Affaires',
      blueprint:  'Blueprint — Graphe du Schéma',
      classes:   'Classes',
      relations: 'Relations Père / Fils',
      functions:  'Fonctions Acteurs',
      namespaces: 'Namespaces',
      templates:  'Templates de Classe',
      mxl:       'Générateur MXL',
      import:    'Import Excel → Classe',
    };

    return { view, ecoName, editClassId, onEcoOpen, navTo, onEditClass, NAV, VIEW_TITLES, toasts };
  },
  template: `
    <div id="app" style="display:flex;height:100vh;overflow:hidden">
      <div id="sidebar">
        <div class="sidebar-logo">
          <div class="sidebar-logo-title">ExoSync Studio</div>
          <div class="sidebar-logo-sub">N1 — Schema Designer</div>
        </div>
        <div class="sidebar-eco" @click="navTo('ecosystem')" :title="ecoName">
          <div style="font-size:0.65rem;text-transform:uppercase;letter-spacing:.08em;margin-bottom:2px">Écosystème actif</div>
          <div class="sidebar-eco-name">{{ ecoName || '—' }}</div>
        </div>
        <nav>
          <div v-for="group in NAV" :key="group.group" class="nav-group">
            <div class="nav-group-label">{{ group.group }}</div>
            <div v-for="item in group.items" :key="item.key"
                 class="nav-item" :class="{active: view===item.key}"
                 @click="navTo(item.key)">
              <span class="nav-icon">{{ item.icon }}</span>
              {{ item.label }}
            </div>
          </div>
        </nav>
        <div class="sidebar-footer">ExoSync Studio N1 · port 8001</div>
      </div>

      <div id="main">
        <div class="topbar">
          <span class="topbar-title">{{ VIEW_TITLES[view] || view }}</span>
          <span class="topbar-spacer"></span>
          <span v-if="view!=='ecosystem'" class="topbar-badge">{{ ecoName }}</span>
        </div>
        <div class="content">
          <view-ecosystem-manager v-if="view==='ecosystem'"  @open="onEcoOpen" />
          <view-workspace          v-if="view==='workspace'" />
          <view-blueprint          v-if="view==='blueprint'" @edit-class="onEditClass" />
          <view-classes            v-if="view==='classes'"   :initial-class-id="editClassId" />
          <view-relations          v-if="view==='relations'" />
          <view-functions          v-if="view==='functions'" />
          <view-namespaces         v-if="view==='namespaces'" />
          <view-templates          v-if="view==='templates'" />
          <view-mxl-preview        v-if="view==='mxl'" />
          <view-import-excel       v-if="view==='import'" />
        </div>
      </div>

      <toast-layer :toasts="toasts" />
    </div>
  `
};

createApp(App).mount('#app');
