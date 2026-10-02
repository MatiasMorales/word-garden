#!/usr/bin/env python3
"""Build index.html from the GRE spreadsheet (Google Sheets "Web page" export).

Usage:  python3 build.py [path/to/GRE.xlsx-folder]

Word definitions live in defs.json and group definitions in groups.json ({"word": "definition"}), edited by hand.
If the spreadsheet export isn't there, the word list already inside index.html is
reused, so definition edits can still be rebuilt.

Each sheet tab (Positivas.html, Negativas.html) is a table where a row with only
column A filled is a group title (its Spanish meaning), followed by word rows:
  A = word, B = part of speech, C = example sentence, D = optional note.
"""
import json, sys
from html.parser import HTMLParser
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else '~/Downloads/GRE.xlsx').expanduser()
HERE = Path(__file__).parent
SHEETS = [('Positivas', 'P'), ('Negativas', 'N')]

# The sheet's group titles and notes are in Spanish; the published game is English-only.
EN_GROUPS = {
    'Abundante': 'Abundant', 'Acuerdo/Consentir': 'Agreement/Consent', 'Adaptable/Flexible': 'Adaptable/Flexible',
    'Agradable de escuchar': 'Pleasant to hear', 'Ahorrativo': 'Thrifty', 'Aprobar/Permitir': 'Approve/Allow',
    'Asistir/Apoyo': 'Assist/Support', 'Audaz/Valiente': 'Bold/Brave', 'Aumentar/Crecer/difundir': 'Increase/Grow/Spread',
    'Autodisciplina / formas de vida': 'Self-discipline/Ways of life', 'Calmado/Sereno': 'Calm/Serene',
    'Calmar/Satisfacer': 'Soothe/Satisfy', 'Caminar/Movimiento': 'Walking/Movement', 'Combinar/Fusionar': 'Combine/Merge',
    'Comienzo/Crecimiento': 'Beginning/Growth', 'Comprender': 'Understand', 'Confirmar/Verificar': 'Confirm/Verify',
    'Curioso/Interesado': 'Curious/Interested', 'Determinación/Firmeza': 'Determination/Firmness',
    'Diligente/Detallista': 'Diligent/Meticulous', 'Diverso/Mixto': 'Diverse/Mixed', 'Elogio/Alabanza': 'Praise/Acclaim',
    'Especular/Adivinar': 'Speculate/Guess', 'Experto/Conocedor': 'Expert/Connoisseur', 'Familia/Parentesco': 'Family/Kinship',
    'Famoso/Destacado': 'Famous/Prominent', 'Fácil de entender': 'Easy to understand', 'Generoso/Amable': 'Generous/Kind',
    'Gracioso/Humor': 'Funny/Humor', 'Grande/Significativo': 'Large/Significant',
    'Igualdad/Imparcialidad': 'Equality/Impartiality', 'Independiente': 'Independent', 'Inteligente/Culto': 'Intelligent/Learned',
    'Investigar/Verificar': 'Investigate/Verify', 'Necesario/Relevante': 'Necessary/Relevant',
    'Neutral/Conectores': 'Neutral/Connectors', 'Nuevo/Original': 'New/Original', 'Optimista': 'Optimistic',
    'Perdonar': 'Forgive', 'Placer/Disfrute': 'Pleasure/Enjoyment', 'Predecir': 'Predict', 'Práctico/Sensato': 'Practical/Sensible',
    'Reducir/Aliviar': 'Reduce/Relieve', 'Sabiduría/Dichos': 'Wisdom/Sayings', 'Saludable/Benéfico': 'Healthy/Beneficial',
    'Sociable': 'Sociable', 'Suerte/Favorable': 'Luck/Favorable', 'Superar/Sobrepasar': 'Overcome/Surpass',
    'Sustituto/Reemplazo': 'Substitute/Replacement', 'Tiempo/Orden/Duración': 'Time/Order/Duration',
    'Uniforme/Similar': 'Uniform/Similar', 'Verdad/Sinceridad': 'Truth/Sincerity',
    'Aburrido / Normal': 'Boring/Ordinary', 'Actuar apresuradamente': 'Acting hastily', 'Anticuado/Obsoleto': 'Outdated/Obsolete',
    'Vergüenza': 'Shame', 'Brechas/Aberturas': 'Gaps/Openings', 'Cambio brusco de humor': 'Sudden mood change',
    'Codicioso': 'Greedy', 'Colapso/Fracaso': 'Collapse/Failure', 'Complejo/Intrincado': 'Complex/Intricate',
    'Comunicación breve': 'Brief communication', 'Criticar': 'Criticize', 'Cruel/Insensible': 'Cruel/Callous',
    'Dañino': 'Harmful', 'Debilitar': 'Weaken', 'Derrochador': 'Wasteful', 'Desacuerdo/Oponerse': 'Disagree/Oppose',
    'Descuidado': 'Careless', 'Dictatorial/Autoritario': 'Dictatorial/Authoritarian', 'Difícil de entender': 'Hard to understand',
    'Dudoso/Escéptico': 'Doubtful/Skeptical', 'Efímero/Temporal': 'Ephemeral/Temporary',
    'Enfermizo/Dañino a la salud': 'Sickly/Unhealthy', 'Escaso': 'Scarce', 'Falso/Imitación': 'Fake/Imitation',
    'Grosero/Verboso': 'Rude/Verbose', 'Hostil': 'Hostile', 'Indeciso/Ambiguo': 'Indecisive/Ambiguous',
    'Indiferente/Apático': 'Indifferent/Apathetic', 'Insincero': 'Insincere', 'Insociable': 'Unsociable',
    'Inusual/Anómalo': 'Unusual/Anomalous', 'Irrelevante': 'Irrelevant', 'Mal humor': 'Bad mood',
    'Mala fama': 'Bad reputation', 'Mentira/Engaño': 'Lie/Deception', 'Mordaz/Cáustico': 'Scathing/Caustic',
    'Muerte/Duelo': 'Death/Mourning', 'Obediente/Sumiso': 'Obedient/Submissive', 'Oculto/Secreto': 'Hidden/Secret',
    'Odio': 'Hatred', 'Parcialidad/Favoritismo': 'Bias/Favoritism', 'Parco/Callado': 'Terse/Quiet',
    'Pequeño/Insignificante': 'Small/Insignificant', 'Perezoso': 'Lazy', 'Pobreza': 'Poverty',
    'Poco práctico/Ingenuo': 'Impractical/Naive', 'Posponer': 'Postpone', 'Prevenir/Obstruir': 'Prevent/Obstruct',
    'Principiante/Inexperto': 'Beginner/Inexperienced', 'Prohibir': 'Prohibit', 'Provocar/Instigar': 'Provoke/Instigate',
    'Regañar/Molestar': 'Scold/Annoy', 'Renunciar/Ceder': 'Give up/Yield', 'Repugnante/Ofensivo': 'Repugnant/Offensive',
    'Retirada/Abandono': 'Withdrawal/Abandonment', 'Separar/Dispersar': 'Separate/Scatter', 'Sin rumbo fijo': 'Aimless',
    'Sonido desagradable': 'Unpleasant sound', 'Terco/Inflexible': 'Stubborn/Inflexible',
    'Timidez/Cobardía': 'Timidity/Cowardice', 'Tristeza/Duelo': 'Sadness/Grief', 'Vacilar/Dudar': 'Waver/Hesitate',
    'Vagar/Divagar': 'Wander/Digress', 'Vicio/Libertinaje': 'Vice/Debauchery',
}
EN_NOTES = {                       # '' drops a note that adds nothing in English
    'pilar': 'pillar', 'sth that cuases th else to change': 'something that causes something else to change',
    'difundir': 'to spread', 'impregnar': 'to soak through', 'dominante': 'widespread', 'cosechar': 'to harvest',
    'abstenerse de alcohol': 'abstaining from alcohol', 'calmed ant controlled': 'calm and controlled',
    '(mas de rendirse)': '(more about giving in)', 'calmar un mal sentimiento': 'to soothe a bad feeling',
    'apaciguar': 'to pacify', 'Falta energía': 'lacking energy', 'Weaknesses': 'a minor weakness',
    'Gaucherio': '', 'Lamentar': 'to regret',
}
# Corrections agreed with the players (2026-10-02)
POS_FIX = {'chide': 'v'}
MERGED = {'deride': 'deride/derisive', 'venerate': 'venerate/veneration'}   # duplicate row -> entry kept
missing = set()


class Table(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self.row, self.cell = [], None, None

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.row = []
        elif tag in ('td', 'th') and self.row is not None:
            self.cell = ''

    def handle_endtag(self, tag):
        if tag in ('td', 'th') and self.cell is not None:
            self.row.append(' '.join(self.cell.split()))
            self.cell = None
        elif tag == 'tr' and self.row is not None:
            self.rows.append(self.row)
            self.row = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell += data


cats, words, index = [], [], {}
for name, code in (SHEETS if SRC.exists() else []):
    t = Table()
    t.feed((SRC / f'{name}.html').read_text(encoding='utf8'))
    cur = None
    for r in t.rows[1:]:                       # first row is the A/B/C/D header
        a, b, c, d = (r + [''] * 5)[1:5]       # column 0 is the row number
        if not (a or b or c):
            continue
        if a and not b and not c:
            if a not in EN_GROUPS:
                missing.add(a)
            cats.append({'n': EN_GROUPS.get(a, a), 's': code})
            cur = len(cats) - 1
            continue
        if cur is None or not a:
            continue
        key = a.lower()
        if key in index:                       # same word listed under two groups
            words[index[key]]['c'].append(cur)
            continue
        index[key] = len(words)
        entry = {'w': a, 'p': b, 'e': c, 'c': [cur]}
        d = EN_NOTES.get(d, d)
        if d and d.lower() not in c.lower():
            entry['t'] = d
        words.append(entry)

if not SRC.exists():
    print(f'{SRC} not found, reusing the word list in index.html')
    old = (HERE / 'index.html').read_text(encoding='utf8')
    start = old.index('const DATA = ') + len('const DATA = ')
    prev = json.loads(old[start:old.index(';\n', start)])
    cats, words = prev['cats'], prev['words']

for entry in words:
    entry['p'] = POS_FIX.get(entry['w'], entry['p'])
words = [e for e in words if e['w'] not in MERGED]

defs = json.loads((HERE / 'defs.json').read_text(encoding='utf8'))
for entry in words:
    entry.pop('d', None)
    if defs.get(entry['w']):
        entry['d'] = defs[entry['w']]
undefined = [e['w'] for e in words if 'd' not in e]
gdefs = json.loads((HERE / 'groups.json').read_text(encoding='utf8'))
for c in cats:
    c.pop('d', None)
    if gdefs.get(c['n']):
        c['d'] = gdefs[c['n']]
undefined += [c['n'] for c in cats if 'd' not in c]

data = json.dumps({'cats': cats, 'words': words, 'merged': MERGED}, ensure_ascii=False, separators=(',', ':'))
html = (HERE / 'template.html').read_text(encoding='utf8').replace('__DATA__', data)
(HERE / 'index.html').write_text(html, encoding='utf8')
print(f'{len(words)} words in {len(cats)} groups -> index.html')
if undefined:
    print(f'⚠️  {len(undefined)} words/groups have no definition in defs.json/groups.json:', ', '.join(undefined[:20]))
if missing:
    print('⚠️  Add English names to EN_GROUPS for:', ', '.join(sorted(missing)))
