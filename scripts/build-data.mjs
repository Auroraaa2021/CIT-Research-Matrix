// Regenerates data.json from the research questionnaire Excel file.
// Usage: node scripts/build-data.mjs [path/to/file.xlsx]
// Requires SheetJS: npm i https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const XLSX_PATH = process.argv[2] || path.join(ROOT, 'data/Research_Questionnaire_Overview_updated-35d04b.xlsx');
const wb = XLSX.read(fs.readFileSync(XLSX_PATH));
const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
const old = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));

const split = s => String(s).split('\n').map(x => x.trim()).filter(x => x && x !== 'None of the above');
const fold = s => s.toLowerCase()
  .replace(/ue/g, 'u').replace(/oe/g, 'o').replace(/ae/g, 'a')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/ß/g, 'ss');
const groupKey = s => fold(s).replace(/\band\b/g, '&').replace(/[^a-z&]/g, '');

// Keep titled names, departments and websites from the existing data where a professor matches.
function findOld(name) {
  const parts = fold(name).split(/\s+/);
  const first = parts[0], last = parts.at(-1);
  const c = old.filter(o => { const on = fold(o.name); return on.includes(first) && on.includes(last); });
  if (c.length === 1) return c[0];
  const byLast = old.filter(o => fold(o.name).split(/\s+/).at(-1) === last);
  return byLast.length === 1 ? byLast[0] : null;
}

const groupDept = new Map();
old.forEach(p => p.research_group.forEach(g => {
  if (g.group_name === 'Neuro-Engineering' && g.group_dept === 'CE') return;
  groupDept.set(groupKey(g.group_name), g.group_dept);
}));
Object.entries({
  'Analysis & PDE': 'MATH',
  'Computer & Communication Systems': 'CE',
  'Security': 'CS',
}).forEach(([k, v]) => groupDept.set(groupKey(k), v));

const IRI_SHORT = s => s.split(' - ')[0].trim();
const OTHER_TUM = {
  'MCQST Excellence Cluster': 'MCQST',
  'MQV: Munich Quantum Valley': 'MQV',
  'MCML: Munich Center for Machine Learning': 'MCML',
  'MUDS: Munich School for Data Science': 'MUDS',
  'BMFTR Research Hub: 6G Life': '6G Life',
  'One Munich Strategy Forum: Next Generation Human-centered Robotics': 'One Munich: Human-centered Robotics',
};
// Free-text answers mapped to normalized unit names; [] means the answer is not a research unit.
const OTHERS = {
  'Max-Plack-Institute for Software Systems': ['Max Planck Institute for Software Systems'],
  'University of Bordeaux': ['University of Bordeaux'],
  'Universidad Complutense de Madrid': ['Universidad Complutense de Madrid'],
  'ZEITlab, ZQE, WSI': ['ZEITlab', 'ZQE', 'WSI'],
  'Konrad Zuse School of Excellence in Reliable AI (relAI)': ['relAI'],
  'relAI (Konrad Zuse School of Excellence in reliable AI)': ['relAI'],
  'relAI': ['relAI'],
  'relAI, METEOR': ['relAI', 'METEOR'],
  'METEOR': ['METEOR'],
  'Max Planck School of Biomedical Artificial Intelligence': ['Max Planck School of Biomedical AI'],
  'ELLIS Munich': ['ELLIS Munich'], 'ELLIS unit Munich': ['ELLIS Munich'], 'ELLIS': ['ELLIS Munich'],
  'M1 Alliance': ['M1 Alliance'],
  'NaKo (AI, MR)': ['NaKo'],
  'Joint research projects with various colleagues in MA, CS, PH, BIO, ASG, LS, MGT, MOS, EE, CE.': [],
  'NFDI (GHGA)': ['NFDI'], 'NFDI (FAIRmat consortium)': ['NFDI'],
  'Center for Digital Medicine and Health': ['Center for Digital Medicine and Health'],
  'EuroHPC': ['EuroHPC'],
  'TUM Georg Nemetschek Institute – AI for the Built World': ['TUM Georg Nemetschek Institute'],
  'ConVeY': ['ConVeY'],
  'BCCN, GSN': ['BCCN', 'GSN'],
  'Bernstein Center for Computational Neuroscience': ['BCCN'],
  'Graduate School for Sytemic Neuroscience': ['GSN'],
  'Research group "Analysis and Mathematical Biology"': [],
  'Graduiertenkolleg Advanced Optimization in a Networked Economy': ['AdONE'],
  'AdONE Research Training Group': ['AdONE'], 'RTG AdONE': ['AdONE'],
  'DFG Research Training Groups AdONE a CONVEY': ['AdONE', 'ConVeY'],
  'Tandem project': [],
  'Jens Grossklags (CIT, Cyber Trust)': [],
  "I'm affiliated with the TUM Campus Straubing for Biotechnology and Sustainability": ['TUM Campus Straubing'],
  'CoC DECS': ['CoC DECS'], 'Design of Electronic Circuits and Systems (DECS)': ['CoC DECS'],
  'CoC-COM': ['CoC COM'],
  'I would like to be affiliated with MIBE': [],
  'MDSI in the process': [],
  'DFG priority program SPP 2265 Random Geometric Systems': ['DFG SPP 2265'],
  'application for a DFG Priority Programme "Higher order interactions"': [],
  'Biosystem Cluster, ZEITlab, COS (Center for Organoid Systems), ARTEMIS Innov. Network, NEUROTECH Innov. Network,':
    ['Biosystem Cluster', 'ZEITlab', 'COS (Center for Organoid Systems)', 'ARTEMIS Innovation Network', 'NEUROTECH Innovation Network'],
  "(I've just started reaching out collaborations with the group of quantum geometry, representation theory (both SoCIT) and the group of philosophy (SOT) - but it's maybe too early to say more.)": [],
  'Joint Global AI Research Hub, Campus Heilbronn': ['Joint Global AI Research Hub Heilbronn'],
  'TUMCREATE, Singapore': ['TUMCREATE Singapore'],
  'TUM Hyperloop, TUMCREATE Singapore': ['TUM Hyperloop', 'TUMCREATE Singapore'],
  'Collaboration with Oliver Hayden for a project he has in TUMCREATE with Singapore': ['TUMCREATE Singapore'],
  'DFG SPP Theoretical Foundations of Deep Learning': ['DFG SPP Theoretical Foundations of Deep Learning'],
  'SFB/TRR 352: Mathematics of Many-Body Quantum Systems and Their Collective Phenomena': ['SFB/TRR 352'],
  '„Mathematics of Many-Body Quantum Systems and Their Collective Phenomena” (SFB TRR352)': ['SFB/TRR 352'],
  'AMC: Atomistic Modelling Center': ['AMC: Atomistic Modelling Center'],
  'TUM Security and Defense Alliance': ['TUM Security and Defense Alliance'],
  'Robotics Institute Germany': ['Robotics Institute Germany'],
  'UPLIFT, DFG Graduiertenschule': ['UPLIFT'],
  'SPP2311, DFG Schwepunkt': ['DFG SPP 2311'],
  'Conemessa (DAAD project)': ['Conemessa (DAAD)'],
  '6G Future Lab Bavaria': ['6G Future Lab Bavaria'],
  'all 6G-life groups, all 6G Bavaria groups': ['6G Future Lab Bavaria'],
  'HDSC/MDSI project Learning and Dynamics in Stochastic Multi-Agent Systems with colleagues from MA (Kuehn, Knochenhauer) nad MGT (Jingui Xie).': [],
};

const unknownOthers = new Set();
const out = rows.filter(r => String(r.Name).trim()).map(r => {
  const excelName = String(r.Name).trim();
  const match = findOld(excelName);

  const research_group = split(r['Active research group']).map(g => ({
    group_name: g,
    group_dept: groupDept.get(groupKey(g)) || 'Unknown',
  }));

  let dept = match?.dept;
  if (!dept) {
    const tally = {};
    research_group.forEach(g => g.group_dept !== 'Unknown' && (tally[g.group_dept] = (tally[g.group_dept] || 0) + 1));
    dept = Object.entries(tally).sort((a, b) => b[1] - a[1])[0]?.[0] || 'CS';
  }

  const inst_type = {};
  split(r['TUM IRIs']).forEach(s => (inst_type[IRI_SHORT(s)] = 'TUM IRI'));
  split(r['Other TUM participated research units']).forEach(s => (inst_type[OTHER_TUM[s] || s] = 'Other TUM'));
  split(r['Other research units']).forEach(s => {
    const mapped = OTHERS[s];
    if (!mapped) { unknownOthers.add(s); inst_type[s] = 'Other'; return; }
    mapped.forEach(m => { if (!inst_type[m]) inst_type[m] = 'Other'; });
  });

  return {
    name: match ? match.name : excelName,
    dept,
    inst: Object.keys(inst_type),
    inst_type,
    research_group,
    'TUM Focus Area': split(r['TUM Focus Area']),
    website: match?.website || '',
    _matched: !!match,
  };
});

console.log('profs', out.length);
console.log('unmatched (dept inferred):', out.filter(p => !p._matched).map(p => `${p.name}[${p.dept}]`).join(', '));
console.log('unknown groups:', [...new Set(out.flatMap(p => p.research_group.filter(g => g.group_dept === 'Unknown').map(g => g.group_name)))]);
console.log('unmapped others:', [...unknownOthers]);
out.forEach(p => delete p._matched);
fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(out, null, 2) + '\n');
