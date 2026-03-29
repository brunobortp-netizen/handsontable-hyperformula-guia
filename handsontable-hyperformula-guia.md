# Guia Completo: Handsontable + HyperFormula (React + TypeScript)

Documentacao tecnica para implementar planilhas Excel-like com suporte a formulas, point-and-click, autocomplete, celulas protegidas, formatacao condicional e auditoria.

---

## 1. Instalacao e Setup

### Pacotes

```bash
npm install handsontable @handsontable/react-wrapper hyperformula
```

### Imports obrigatorios

```tsx
import { HotTable } from '@handsontable/react-wrapper'
import type { HotTableRef } from '@handsontable/react-wrapper'
import { registerAllModules } from 'handsontable/registry'
import { HyperFormula } from 'hyperformula'
import Handsontable from 'handsontable'
import 'handsontable/styles/handsontable.min.css'
import 'handsontable/styles/ht-theme-main.min.css'
```

**CRITICO:** Chamar `registerAllModules()` no top-level do modulo (fora do componente):

```tsx
registerAllModules()
```

---

## 2. HyperFormula com Idioma Localizado (ex: Portugues)

### Problema: erro "Language not registered"

O Handsontable sobrescreve a configuracao de idioma do HyperFormula para `enGB` por padrao. Se voce criar uma instancia pre-configurada do HyperFormula com `ptPT` e passa-la ao Handsontable, ele IGNORA o idioma e usa `enGB`, causando `#NAME?` em formulas como `=SOMA()`.

### Solucao correta

1. Registrar o idioma ANTES de tudo (top-level):

```tsx
import { ptPT } from 'hyperformula/i18n/languages'
HyperFormula.registerLanguage('ptPT', ptPT)
```

2. Passar a CLASSE HyperFormula (nao uma instancia) dentro de `formulas.engine`:

```tsx
const formulasConfig = useMemo(() => ({
  engine: {
    hyperformula: HyperFormula,    // CLASSE, nao instancia!
    language: 'ptPT',
    functionArgSeparator: ';',
    decimalSeparator: ',' as const,
    thousandSeparator: '.' as const,
  },
  language: ptPT,
  sheetName: 'MinhaPlanilha',
}), [])
```

### O que NAO fazer

```tsx
// ERRADO — Handsontable ignora o idioma da instancia
const hfInstance = HyperFormula.buildEmpty({ language: 'ptPT' })
// ... formulas={{ engine: hfInstance }} => #NAME? em SOMA, MEDIA, etc.

// ERRADO — passar instancia pre-construida
const hfInstance = HyperFormula.buildEmpty({ ... })
// formulas={{ engine: hfInstance }}

// CORRETO — passar a classe dentro de engine config
// formulas={{ engine: { hyperformula: HyperFormula, language: 'ptPT', ... } }}
```

### Idiomas disponiveis

O HyperFormula traz locales em `hyperformula/es/i18n/languages/`. Verifique os arquivos disponiveis:

```
enGB.mjs, ptPT.mjs, frFR.mjs, deDE.mjs, esES.mjs, plPL.mjs, csCZ.mjs
```

Cada locale exporta um dicionario mapeando nomes de funcoes. Ex: `ptPT` mapeia `SUM` -> `SOMA`, `AVERAGE` -> `MEDIA`, `IF` -> `SE`, etc.

---

## 3. Dados e Formulas

### Estrutura de dados

Os dados sao uma matriz 2D (`(string | number | null)[][]`). NAO inclua o header como primeira linha dos dados — use `colHeaders` para isso.

```tsx
const spreadsheetData = useMemo(() => {
  const data: (string | number | null)[][] = []

  // Linhas de dados (row 0 no array = row 1 no HyperFormula)
  for (let i = 0; i < items.length; i++) {
    const rowNum = i + 1  // HyperFormula usa 1-indexed nas referencias
    data.push([
      items[i].label,
      items[i].value1,
      items[i].value2,
      items[i].value3,
      `=SOMA(B${rowNum}:D${rowNum})`,  // Total horizontal
    ])
  }

  // Linha de totais (formulas verticais)
  const lastRow = items.length
  data.push([
    'Total',
    `=SOMA(B1:B${lastRow})`,
    `=SOMA(C1:C${lastRow})`,
    `=SOMA(D1:D${lastRow})`,
    `=SOMA(E1:E${lastRow})`,
  ])

  return data
}, [items])
```

**CRITICO — Indexacao:** Os dados no array sao 0-indexed, mas as referencias de celulas do HyperFormula sao 1-indexed. `data[0]` corresponde a row 1 nas formulas. Se voce incluir o header como row 0 dos dados, TODAS as formulas ficam deslocadas e dao `#NAME?` ou valores errados.

---

## 4. Headers de Linha e Coluna

### Mostrar referencia da coluna (A, B, C...) acima do nome

Para que o usuario saiba montar formulas, exiba a letra da coluna:

```tsx
function colName(col: number): string {
  return String.fromCharCode(65 + col)
}

const colHeaderNames = ['Categoria', 'Jan', 'Fev', 'Mar', 'Total']
const colHeaders = colHeaderNames.map((name, i) =>
  `<div style="line-height:1.2">
    <span style="font-size:10px;color:var(--color-text-secondary);font-weight:400">
      ${colName(i)}
    </span>
    <br/>${name}
  </div>`
)
```

### Numeros de linha nativos

```tsx
<HotTable rowHeaders={true} ... />
```

Isso exibe 1, 2, 3... automaticamente, consistente com as referencias do HyperFormula.

---

## 5. Celulas Protegidas (Read-Only)

Use a funcao `cells` para definir quais celulas sao editaveis:

```tsx
const cellsFunction = useCallback(function (
  this: Handsontable.CellProperties, row: number, col: number
): Handsontable.CellMeta {
  const meta: Handsontable.CellMeta = {}
  ;(meta as any).renderer = customRendererFn

  const hot = hotRef.current?.hotInstance
  const totalRowIdx = hot ? hot.countRows() - 1 : 6

  // Proteger: coluna de categorias (0), coluna de total (4), linha de total
  if (col === 0 || col === 4 || row === totalRowIdx) {
    meta.editor = false as unknown as string
  }

  return meta
}, [customRendererFn])
```

**Nota:** `meta.editor = false` impede a edicao. O cast `as unknown as string` contorna o tipo TypeScript.

---

## 6. Custom Renderer (Formatacao Condicional)

O renderer customizado aplica classes CSS baseadas no valor da celula:

```tsx
const customRendererFn = useCallback(function (
  this: unknown,
  instance: Handsontable,
  td: HTMLTableCellElement,
  row: number, col: number,
  prop: string | number,
  value: unknown,
  cellProperties: Handsontable.CellProperties
) {
  // SEMPRE chamar o renderer base primeiro
  Handsontable.renderers.TextRenderer.call(
    this, instance, td, row, col, prop, value, cellProperties
  )

  // Limpar classes anteriores para evitar acumulo
  td.classList.remove(
    'negative-value', 'zero-value', 'over-budget',
    'formula-cell', 'total-row-cell', 'category-cell',
    'read-only-cell', 'ref-highlight'
  )

  // Highlight visual para celulas referenciadas (point-and-click)
  const hl = highlightedRangeRef.current
  if (hl && row >= hl.r1 && row <= hl.r2 && col >= hl.c1 && col <= hl.c2) {
    td.classList.add('ref-highlight')
  }

  const totalRowIdx = instance.countRows() - 1

  // Indicador de formula (triangulo no canto)
  const src = instance.getSourceDataAtCell(row, col)
  if (typeof src === 'string' && src.startsWith('=')) {
    td.classList.add('formula-cell')
    td.style.position = 'relative'  // necessario para o ::after
  }

  // Formatacao numerica
  const numVal = typeof value === 'number' ? value : parseFloat(String(value))
  if (!isNaN(numVal) && col >= 1) {
    if (numVal < 0) td.classList.add('negative-value')
    else if (numVal === 0) td.classList.add('zero-value')

    // formatacao condicional aplicada via classes CSS
  }
}, [])
```

**CRITICO:** Use `highlightedRangeRef.current` (ref) dentro do renderer, NAO o state diretamente. O renderer e chamado pelo Handsontable fora do ciclo de render do React — closures sobre state ficam stale.

---

## 7. Formula Bar

Barra acima da planilha que mostra a referencia da celula selecionada (ex: `B3`) e a formula/valor da celula.

### Estado necessario

```tsx
const [selectedCell, setSelectedCell] = useState<string>('A1')
const [formulaBarValue, setFormulaBarValue] = useState<string>('')
const [isEditingFormulaBar, setIsEditingFormulaBar] = useState(false)
const formulaInputRef = useRef<HTMLInputElement>(null)
const formulaBarValueRef = useRef(formulaBarValue)
formulaBarValueRef.current = formulaBarValue
```

### Sincronizar com selecao

```tsx
function cellRef(row: number, col: number): string {
  return `${colName(col)}${row + 1}`
}

const handleAfterSelection = useCallback((row: number, col: number) => {
  const hot = hotRef.current?.hotInstance
  if (!hot) return
  setSelectedCell(cellRef(row, col))
  const src = hot.getSourceDataAtCell(row, col)
  setFormulaBarValue(src != null ? String(src) : '')
  currentRow.current = row
  currentCol.current = col
}, [])
```

**IMPORTANTE:** Use `getSourceDataAtCell` para obter a formula original (ex: `=SOMA(B1:D1)`), nao `getDataAtCell` que retorna o valor calculado.

### Confirmar e cancelar edicao

```tsx
const confirmFormulaBar = useCallback(() => {
  const hot = hotRef.current?.hotInstance
  if (!hot) return
  if (currentRow.current >= 0 && currentCol.current >= 0) {
    hot.setDataAtCell(currentRow.current, currentCol.current, formulaBarValue)
  }
  hideAutocomplete()
  clearNavigation()
  isFormulaMode.current = false
}, [formulaBarValue, hideAutocomplete, clearNavigation])
```

---

## 8. Autocomplete de Formulas

Dropdown que aparece ao digitar `=SO...` sugerindo funcoes como `SOMA`, `SOMASE`.

### Dicionario completo de funcoes (ptPT)

O dicionario DEVE conter TODAS as funcoes que o usuario pode precisar. Se `SOMASE` nao estiver no dicionario, digitar `=SOM` mostra apenas `SOMA` — o usuario nao descobre que `SOMASE` existe.

```tsx
const FORMULA_DICT: Record<string, string> = {
  // Matematica e Agregacao
  'SOMA': 'SOMA(intervalo)',
  'SOMASE': 'SOMASE(intervalo_criterio; criterio; intervalo_soma)',
  'SOMASES': 'SOMASES(intervalo_soma; intervalo1; criterio1; ...)',
  'SOMARPRODUTO': 'SOMARPRODUTO(matriz1; matriz2; ...)',
  'SOMAQUAD': 'SOMAQUAD(intervalo)',
  'MULT': 'MULT(intervalo)',
  'MEDIA': 'MEDIA(intervalo)',
  'MEDIASE': 'MEDIASE(intervalo_criterio; criterio; intervalo_media)',
  'MED': 'MED(intervalo)',
  'MAXIMO': 'MAXIMO(intervalo)',
  'MAXIMOSES': 'MAXIMOSES(intervalo_max; intervalo1; criterio1; ...)',
  'MINIMO': 'MINIMO(intervalo)',
  'MINIMOSES': 'MINIMOSES(intervalo_min; intervalo1; criterio1; ...)',
  'CONT.NUM': 'CONT.NUM(intervalo)',
  'CONT.SE': 'CONT.SE(intervalo; criterio)',
  'CONT.SES': 'CONT.SES(intervalo1; criterio1; intervalo2; criterio2; ...)',
  'CONT.VALORES': 'CONT.VALORES(intervalo)',
  'CONTAR.VAZIO': 'CONTAR.VAZIO(intervalo)',
  'SUBTOTAL': 'SUBTOTAL(funcao; intervalo)',
  // Matematica Basica
  'ABS': 'ABS(numero)',
  'ARRED': 'ARRED(numero; casas)',
  'ARREDONDAR.PARA.BAIXO': 'ARREDONDAR.PARA.BAIXO(numero; casas)',
  'ARREDONDAR.PARA.CIMA': 'ARREDONDAR.PARA.CIMA(numero; casas)',
  'INT': 'INT(numero)',
  'MOD': 'MOD(numero; divisor)',
  'POTENCIA': 'POTENCIA(base; expoente)',
  'RAIZ': 'RAIZ(numero)',
  'PI': 'PI()',
  'TRUNCAR': 'TRUNCAR(numero; casas)',
  'SINAL': 'SINAL(numero)',
  'QUOCIENTE': 'QUOCIENTE(numerador; denominador)',
  'PAR': 'PAR(numero)',
  'IMPAR': 'IMPAR(numero)',
  // Logica
  'SE': 'SE(condicao; verdadeiro; falso)',
  'SE.S': 'SE.S(condicao1; valor1; condicao2; valor2; ...)',
  'SEERRO': 'SEERRO(valor; valor_se_erro)',
  'SENA': 'SENA(valor; valor_se_na)',
  'NAO': 'NAO(logico)',
  'OU': 'OU(logico1; logico2; ...)',
  'OUEXCL': 'OUEXCL(logico1; logico2; ...)',
  'VERDADEIRO': 'VERDADEIRO()',
  // Lookup e Referencia
  'PROCV': 'PROCV(valor; tabela; coluna; correspondencia)',
  'PROCH': 'PROCH(valor; tabela; linha; correspondencia)',
  'PROCX': 'PROCX(valor; intervalo_busca; intervalo_resultado)',
  'INDICE': 'INDICE(matriz; linha; coluna)',
  'CORRESP': 'CORRESP(valor; intervalo; tipo)',
  'DESLOC': 'DESLOC(ref; linhas; colunas; altura; largura)',
  'LIN': 'LIN(referencia)',
  'COL': 'COL(referencia)',
  'LINS': 'LINS(intervalo)',
  'COLS': 'COLS(intervalo)',
  // Texto
  'CONCATENAR': 'CONCATENAR(texto1; texto2; ...)',
  'TEXTO': 'TEXTO(valor; formato)',
  'VALOR': 'VALOR(texto)',
  'MAIUSCULA': 'MAIUSCULA(texto)',
  'MINUSCULA': 'MINUSCULA(texto)',
  'PRI.MAIUSCULA': 'PRI.MAIUSCULA(texto)',
  'ARRUMAR': 'ARRUMAR(texto)',
  'SUBSTITUIR': 'SUBSTITUIR(texto; antigo; novo; instancia)',
  'NUM.CARACT': 'NUM.CARACT(texto)',
  'DIREITA': 'DIREITA(texto; num_caract)',
  // Data
  'AGORA': 'AGORA()',
  'DATA': 'DATA(ano; mes; dia)',
  'ANO': 'ANO(data)',
  'MES': 'MES(data)',
  'DIA': 'DIA(data)',
  'DIAS': 'DIAS(data_final; data_inicial)',
  // Estatistica
  'DESVPAD': 'DESVPAD(intervalo)',
  'VAR': 'VAR(intervalo)',
  'MAIOR': 'MAIOR(intervalo; k)',
  'MENOR': 'MENOR(intervalo; k)',
  'CORREL': 'CORREL(intervalo1; intervalo2)',
}
```

> **REGRA:** Sempre inclua as variantes condicionais (`SOMASE`, `CONT.SE`, `MEDIASE`, `MAXIMOSES`, `MINIMOSES`, `SOMASES`, `CONT.SES`). O usuario que digita `=SOM` espera ver `SOMA` E `SOMASE` no dropdown.

### Logica de matching

```tsx
const showAutocomplete = useCallback((val: string, anchor) => {
  // Match apos =, +, -, *, /, ;, (
  const match = val.match(/(?:^=|[+\-*/;(])([A-Za-z.]+)$/)
  if (!match) { setAcItems([]); setAcPos(null); return }
  const typed = match[1].toUpperCase()
  const filtered = Object.entries(FORMULA_DICT)
    .filter(([name]) => name.startsWith(typed))
    .map(([name, sig]) => ({ name, sig }))
  if (filtered.length === 0) { setAcItems([]); setAcPos(null); return }
  setAcItems(filtered)
  setAcIdx(0)
  setAcPos(anchor)
}, [])
```

### Inserir funcao selecionada

```tsx
const insertAutocompleteFn = useCallback((funcName: string, currentVal: string): string => {
  const match = currentVal.match(/([A-Za-z.]+)$/)
  if (match) {
    return currentVal.slice(0, currentVal.length - match[1].length) + funcName + '('
  }
  return currentVal + funcName + '('
}, [])
```

### Funcionar no editor inline da celula (CRITICO)

O Handsontable usa um `<textarea>` interno como editor. Para o autocomplete funcionar tanto na formula bar quanto no editor inline, voce DEVE:

1. No hook `afterBeginEditing`, capturar o textarea do editor ativo
2. Adicionar listeners de `input` e `keydown` no textarea
3. Usar refs para o state do autocomplete (evitar stale closures)
4. Limpar listeners no `afterDeselect`

```tsx
const cellEditorInputRef = useRef<HTMLTextAreaElement | null>(null)
const acItemsRef = useRef(acItems)
const acIdxRef = useRef(acIdx)
acItemsRef.current = acItems
acIdxRef.current = acIdx

const handleAfterBeginEditing = useCallback((row: number, col: number) => {
  const hot = hotRef.current?.hotInstance
  if (!hot) return
  const editor = hot.getActiveEditor()
  if (!editor) return
  const textarea = (editor as any).TEXTAREA as HTMLTextAreaElement | undefined
  if (!textarea) return
  cellEditorInputRef.current = textarea

  const onInput = () => {
    if (isNavigatingRef.current) clearNavigation()
    forceTextModeRef.current = false
    const val = textarea.value
    isFormulaMode.current = val.startsWith('=')
    setFormulaBarValue(val)
    if (val.startsWith('=')) {
      const rect = textarea.getBoundingClientRect()
      showAutocomplete(val, { top: rect.bottom + 2, left: rect.left })
    } else {
      hideAutocomplete()
    }
  }

  const onKeyDown = (e: KeyboardEvent) => {
    const items = acItemsRef.current  // usar REF, nao state
    const idx = acIdxRef.current
    if (items.length > 0) {
      if (e.key === 'ArrowDown') {
        e.preventDefault(); e.stopPropagation()
        setAcIdx(i => Math.min(i + 1, items.length - 1)); return
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault(); e.stopPropagation()
        setAcIdx(i => Math.max(i - 1, 0)); return
      }
      if (e.key === 'Tab') {
        e.preventDefault(); e.stopPropagation()
        const newVal = insertAutocompleteFn(items[idx].name, textarea.value)
        textarea.value = newVal
        setFormulaBarValue(newVal)
        hideAutocomplete()
        textarea.focus()
        textarea.selectionStart = textarea.selectionEnd = newVal.length
        return
      }
      if (e.key === 'Escape') { e.preventDefault(); hideAutocomplete(); return }
    }
  }

  textarea.addEventListener('input', onInput)
  textarea.addEventListener('keydown', onKeyDown as EventListener)
  cellEditorListenerRef.current = onInput
  cellEditorKeyListenerRef.current = onKeyDown

  isEditingInline.current = true

  const src = hot.getSourceDataAtCell(row, col)
  if (typeof src === 'string' && src.startsWith('=')) {
    isFormulaMode.current = true
  }
}, [showAutocomplete, hideAutocomplete, insertAutocompleteFn, clearNavigation])
```

**REGRA:** Quando o autocomplete esta aberto, as setas ArrowUp/ArrowDown DEVEM navegar o dropdown, NAO as celulas. Verificar `acItemsRef.current.length > 0` antes de qualquer logica de navegacao de celulas.

---

## 9. Point-and-Click para Formulas (Comportamento Excel-like)

Este e o sistema mais complexo. Permite que o usuario digite `=SOMA(` e clique em celulas ou use setas para inserir referencias.

### 9.1 Estado de navegacao

```tsx
const isFormulaMode = useRef(false)         // editando celula com "="
const isEditingInline = useRef(false)        // editando na celula (vs formula bar)
const currentRow = useRef(-1)
const currentCol = useRef(-1)

const isNavigatingRef = useRef(false)        // setas/cliques movendo cursor de ref
const forceTextModeRef = useRef(false)       // F2 toggle: forca modo texto
const refAnchor = useRef({ row: 0, col: 0 })
const refCursorRef = useRef({ row: 0, col: 0 })
const refInsertPos = useRef(0)               // posicao no texto onde ref comeca
const refSuffix = useRef('')                 // texto apos cursor (preservar mid-text)

const isDraggingRef = useRef(false)
const dragStartRef = useRef({ row: 0, col: 0 })

const [highlightedRange, setHighlightedRange] = useState<{
  r1: number; c1: number; r2: number; c2: number
} | null>(null)
const highlightedRangeRef = useRef(highlightedRange)
highlightedRangeRef.current = highlightedRange
```

### 9.2 Detectar ponto de insercao de referencia (CRITICO)

Esta funcao decide se as setas devem **mover o cursor de texto** ou **navegar celulas e inserir referencia**.

```tsx
const REF_INSERT_CHARS = '(;+-*/=,'

function isAtRefInsertionPoint(text: string, cursorPos: number): boolean {
  if (!text.startsWith('=')) return false
  if (cursorPos <= 0) return false
  if (cursorPos === 1) return true  // logo apos o "="

  const charBefore = text[cursorPos - 1]

  // Cursor no FIM do texto
  if (cursorPos >= text.length) {
    return REF_INSERT_CHARS.includes(charBefore)
  }

  // Cursor no MEIO: so ativa se char antes E char depois sao delimitadores
  // Pega o caso "slot vazio": =SE(A1>0;|;0) onde usuario apagou argumento
  if (REF_INSERT_CHARS.includes(charBefore)) {
    const charAfter = text[cursorPos]
    if (charAfter === ';' || charAfter === ')' || charAfter === ',') return true
  }

  return false
}
```

**Por que essa logica e importante:**

| Cenario | charBefore | charAfter | Resultado |
|---------|------------|-----------|-----------|
| `=SOMA(\|` (fim) | `(` | - | NAVEGAR celulas |
| `=SOMA(B1\|` (fim) | `1` | - | mover texto |
| `=SE(A1>0;\|;0)` (meio) | `;` | `;` | NAVEGAR celulas |
| `=SE(A1>0;1\|00;0)` (meio) | `1` | `0` | mover texto |

**Erro comum:** Fazer `if (cursorPos < text.length) return false` indiscriminadamente impede insercao mid-text.

**Outro erro:** Aceitar qualquer `REF_INSERT_CHARS` antes do cursor sem verificar o que vem depois faz com que ao navegar com setas, a seta pegue referencia em vez de mover o cursor.

### 9.3 Ler e escrever texto da formula (DOM-first, NUNCA React state)

**REGRA CRITICA:** `getFormulaText()` e `getCursorPos()` devem SEMPRE ler do elemento DOM ativo (`textarea.value` / `textarea.selectionStart`), **nunca** de React state ou refs sincronizados via `useEffect`.

**Por que:** O React state e assIncrono. Entre o usuario digitar `;` (evento `input` -> `setState`) e apertar seta (evento `beforeKeyDown`), o React pode NAO ter renderizado ainda. O ref/state contem o valor antigo sem o `;`. Isso faz `isAtRefInsertionPoint` retornar `false` quando deveria retornar `true`, quebrando point-and-click apos digitar operadores/separadores.

**Hierarquia de leitura (mesma para getFormulaText e getCursorPos):**
1. Se editor inline esta aberto (`isEditingInline.current && cellEditorInputRef.current`) -> ler de `cellEditorInputRef.current`
2. Se formula bar tem foco (`formulaInputRef.current` focado) -> ler de `formulaInputRef.current`
3. Fallback -> ref sincronizado (`formulaBarValueRef.current`)

```tsx
// CORRETO — le SEMPRE do DOM, nunca de React state
const getFormulaText = useCallback(() => {
  if (isEditingInline.current && cellEditorInputRef.current) {
    return cellEditorInputRef.current.value  // DOM direto
  }
  if (formulaInputRef.current && document.activeElement === formulaInputRef.current) {
    return formulaInputRef.current.value     // DOM direto
  }
  return formulaBarValueRef.current          // fallback
}, [])

// ERRADO — React state pode estar desatualizado no beforeKeyDown
// const getFormulaText = () => formulaBarValue   // NUNCA
// const getFormulaText = () => someRef.current   // ref via useEffect tambem falha
```

```tsx
// IMPORTANTE: checar isEditingInline PRIMEIRO, sem depender
// de document.activeElement. Dentro do beforeKeyDown do
// Handsontable, document.activeElement NAO e confiavel —
// o sistema de eventos intercepta o evento antes que o
// browser atualize o foco. Se depender de activeElement,
// getCursorPos retorna o fallback (text.length) e
// isAtRefInsertionPoint avalia com posicao errada —
// causando falha silenciosa no point-and-click
// (ex: =SOMA( + seta move cursor em vez de inserir ref).
const getCursorPos = useCallback((): number => {
  if (isEditingInline.current && cellEditorInputRef.current) {
    return cellEditorInputRef.current.selectionStart
      ?? cellEditorInputRef.current.value.length
  }
  if (formulaInputRef.current
    && document.activeElement === formulaInputRef.current) {
    return formulaInputRef.current.selectionStart
      ?? formulaInputRef.current.value.length
  }
  return formulaBarValueRef.current.length
}, [])
```

```tsx
const setFormulaText = useCallback((text: string) => {
  if (isEditingInline.current && cellEditorInputRef.current) {
    const textarea = cellEditorInputRef.current
    textarea.value = text
    textarea.selectionStart = textarea.selectionEnd = text.length
    // NAO dispatch 'input' — causaria loop e resetaria navegacao
  }
  setFormulaBarValue(text)
}, [])

const clearNavigation = useCallback(() => {
  isNavigatingRef.current = false
  forceTextModeRef.current = false
  refSuffix.current = ''
  setHighlightedRange(null)
}, [])

const updateNavRef = useCallback((
  anchorRow: number, anchorCol: number,
  cursorRow: number, cursorCol: number
) => {
  const text = getFormulaText()

  if (!isNavigatingRef.current) {
    const cursorPos = getCursorPos()
    refInsertPos.current = cursorPos
    refSuffix.current = text.slice(cursorPos)  // CRITICO
    isNavigatingRef.current = true
  }

  refAnchor.current = { row: anchorRow, col: anchorCol }
  refCursorRef.current = { row: cursorRow, col: cursorCol }

  const r1 = Math.min(anchorRow, cursorRow)
  const c1 = Math.min(anchorCol, cursorCol)
  const r2 = Math.max(anchorRow, cursorRow)
  const c2 = Math.max(anchorCol, cursorCol)

  const ref = (r1 === r2 && c1 === c2)
    ? cellRef(r1, c1)
    : `${cellRef(r1, c1)}:${cellRef(r2, c2)}`

  const newText = text.slice(0, refInsertPos.current) + ref + refSuffix.current
  setFormulaText(newText)
  setHighlightedRange({ r1, c1, r2, c2 })

  // Reposicionar cursor logo apos a referencia inserida
  const cursorTarget = refInsertPos.current + ref.length
  requestAnimationFrame(() => {
    if (isEditingInline.current && cellEditorInputRef.current) {
      cellEditorInputRef.current.selectionStart = cursorTarget
      cellEditorInputRef.current.selectionEnd = cursorTarget
    } else if (formulaInputRef.current) {
      formulaInputRef.current.selectionStart = cursorTarget
      formulaInputRef.current.selectionEnd = cursorTarget
    }
  })
}, [getFormulaText, setFormulaText, getCursorPos])
```

**CRITICO — Preservar sufixo (refSuffix):**

Sem o `refSuffix`, ao inserir referencia no meio de `=SE(A1>0;|;0)`:
- ERRADO: `=SE(A1>0;D4` (perdeu `;0)`)
- CORRETO: `=SE(A1>0;D4;0)` (preservou `;0)`)

**CRITICO — Reposicionar cursor apos inserir referencia:**

Apos `setFormulaText(newText)`, o cursor deve ficar logo apos a referencia inserida (antes do sufixo). Sem isso, o cursor fica no fim do texto ou em posicao errada, fazendo a proxima operacao falhar (Shift+Arrow para expandir range, ou digitar `;` para proximo argumento).

**Por que `requestAnimationFrame`:** O `setFormulaText` altera o `value` do textarea, mas o React pode re-renderizar a formula bar (via `setFormulaBarValue`) e resetar a posicao. O rAF garante que o reposicionamento acontece APOS o render.

**Posicao correta:** `refInsertPos + ref.length` = logo apos a referencia, antes do sufixo. Ex: `=SE(A1>0;|D4|;0)` — cursor entre D4 e ;0).

### 9.4 Interceptar clique em celula

```tsx
const handleBeforeMouseDown = useCallback((
  event: MouseEvent, coords: Handsontable.CellCoords
) => {
  if (!isFormulaMode.current) return
  if (coords.row < 0 || coords.col < 0) return

  if (!isNavigatingRef.current && !forceTextModeRef.current) {
    const text = getFormulaText()
    const cursorPos = getCursorPos()
    if (!isAtRefInsertionPoint(text, cursorPos)) {
      clearNavigation()
      isFormulaMode.current = false
      isEditingInline.current = false
      return  // clique confirma edicao
    }
  }
  if (forceTextModeRef.current) {
    clearNavigation()
    isFormulaMode.current = false
    isEditingInline.current = false
    return
  }

  event.stopImmediatePropagation()
  event.preventDefault()

  isDraggingRef.current = true
  dragStartRef.current = { row: coords.row, col: coords.col }

  if (event.shiftKey && isNavigatingRef.current) {
    // Shift+click: estender range
    updateNavRef(refAnchor.current.row, refAnchor.current.col,
      coords.row, coords.col)
    isDraggingRef.current = false
  } else if (event.ctrlKey || event.metaKey) {
    // Ctrl+click: adicionar ";" e nova ref
    if (isNavigatingRef.current) {
      const text = getFormulaText()
      const lastChar = text.slice(-1)
      if (lastChar && /[A-Za-z0-9]/.test(lastChar)) {
        setFormulaText(text + ';')
      }
      isNavigatingRef.current = false
    }
    updateNavRef(coords.row, coords.col, coords.row, coords.col)
  } else {
    updateNavRef(coords.row, coords.col, coords.row, coords.col)
  }

  // Manter foco no editor
  if (isEditingInline.current && cellEditorInputRef.current) {
    cellEditorInputRef.current.focus()
  } else {
    formulaInputRef.current?.focus()
  }
}, [updateNavRef, getFormulaText, getCursorPos,
  setFormulaText, clearNavigation])
```

### 9.5 Drag para selecao de range

```tsx
useEffect(() => {
  const onUp = () => { isDraggingRef.current = false }
  document.addEventListener('mouseup', onUp)
  return () => document.removeEventListener('mouseup', onUp)
}, [])

const handleBeforeMouseOver = useCallback((
  event: MouseEvent, coords: Handsontable.CellCoords
) => {
  if (!isFormulaMode.current || !isDraggingRef.current) return
  if (coords.row < 0 || coords.col < 0) return

  event.stopImmediatePropagation()
  event.preventDefault()

  updateNavRef(dragStartRef.current.row, dragStartRef.current.col,
    coords.row, coords.col)

  if (isEditingInline.current && cellEditorInputRef.current) {
    cellEditorInputRef.current.focus()
  } else {
    formulaInputRef.current?.focus()
  }
}, [updateNavRef])
```

### 9.6 Interceptar teclado (beforeKeyDown)

```tsx
const handleBeforeKeyDown = useCallback((event: KeyboardEvent) => {
  if (!isFormulaMode.current) return

  if (event.key === 'Escape') {
    clearNavigation(); isFormulaMode.current = false
    isEditingInline.current = false; return
  }
  if (event.key === 'Enter') {
    clearNavigation(); isFormulaMode.current = false
    isEditingInline.current = false; return
  }
  if (event.key === 'F2') {
    event.preventDefault(); event.stopImmediatePropagation()
    if (isNavigatingRef.current) clearNavigation()
    forceTextModeRef.current = !forceTextModeRef.current
    return
  }

  if (!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown']
    .includes(event.key)) return
  if (event.ctrlKey || event.metaKey) return
  if (acItemsRef.current.length > 0) return

  const hot = hotRef.current?.hotInstance
  if (!hot) return

  const editor = hot.getActiveEditor()
  const editorOpen = editor && (editor as any).isOpened?.()
  const formulaBarFocused =
    document.activeElement === formulaInputRef.current
  if (!editorOpen && !formulaBarFocused) return

  if (!isNavigatingRef.current) {
    if (forceTextModeRef.current) return
    const text = getFormulaText()
    const cursorPos = getCursorPos()
    if (!isAtRefInsertionPoint(text, cursorPos)) return
  }

  event.preventDefault()
  event.stopImmediatePropagation()
  ;(event as any).isImmediatePropagationStopped = () => true

  const maxRow = hot.countRows() - 1
  const maxCol = hot.countCols() - 1

  let r = isNavigatingRef.current
    ? refCursorRef.current.row : currentRow.current
  let c = isNavigatingRef.current
    ? refCursorRef.current.col : currentCol.current

  if (event.key === 'ArrowLeft') c = Math.max(0, c - 1)
  else if (event.key === 'ArrowRight') c = Math.min(maxCol, c + 1)
  else if (event.key === 'ArrowUp') r = Math.max(0, r - 1)
  else if (event.key === 'ArrowDown') r = Math.min(maxRow, r + 1)

  if (event.shiftKey) {
    if (!isNavigatingRef.current) {
      refInsertPos.current = getFormulaText().length
      isNavigatingRef.current = true
      refAnchor.current = { row: r, col: c }
    }
    updateNavRef(refAnchor.current.row,
      refAnchor.current.col, r, c)
  } else {
    updateNavRef(r, c, r, c)
  }
}, [getFormulaText, getCursorPos, updateNavRef, clearNavigation])
```

**CRITICO:** `(event as any).isImmediatePropagationStopped = () => true` — necessario porque o Handsontable verifica essa funcao internamente.

### 9.7 Resetar navegacao ao digitar

```tsx
// No handler de input do editor inline
const onInput = () => {
  if (isNavigatingRef.current) clearNavigation()
  forceTextModeRef.current = false
  // ...
}

// No handler de change da formula bar
const handleFormulaBarChange = useCallback((val: string) => {
  if (isNavigatingRef.current) clearNavigation()
  forceTextModeRef.current = false
  // ...
}, [])
```

---

## 10. Props do HotTable (Configuracao Completa)

```tsx
<HotTable
  ref={hotRef}
  data={spreadsheetData}
  colHeaders={colHeaders}
  rowHeaders={true}
  width="100%"
  height="auto"
  stretchH="all"
  licenseKey="non-commercial-and-evaluation"
  formulas={formulasConfig}
  contextMenu={true}
  undo={true}
  copyPaste={{ pasteMode: 'overwrite' }}
  fillHandle={true}
  manualColumnResize={true}
  afterSelection={handleAfterSelection}
  afterChange={handleAfterChange}
  beforeOnCellMouseDown={handleBeforeMouseDown}
  beforeOnCellMouseOver={handleBeforeMouseOver}
  afterBeginEditing={handleAfterBeginEditing}
  afterDeselect={handleAfterDeselect}
  beforeKeyDown={handleBeforeKeyDown}
  cells={cellsFunction}
  className="ht-theme-main"
/>
```

| Prop | Valor | Por que |
|------|-------|---------|
| `height` | `"auto"` | `"100%"` pode causar tela em branco |
| `fillHandle` | `true` | Arrastar em TODAS as direcoes. `{ direction: 'vertical' }` restringe |
| `formulas` | config object | Passar CLASSE, nao instancia |
| `stretchH` | `"all"` | Estica colunas para preencher largura |

---

## 11. CSS Necessario

```css
/* Container */
.ht-spreadsheet .handsontable {
  font-family: 'Inter', sans-serif;
  font-size: 13px;
}

/* Headers */
.ht-spreadsheet .handsontable th {
  background-color: #f1f5f9;
  font-weight: 600;
  font-size: 12px;
  border-color: var(--color-border) !important;
}

/* Celulas */
.ht-spreadsheet .handsontable td {
  border-color: var(--color-border) !important;
  background-color: var(--color-surface);
}

/* Protegida */
.ht-spreadsheet .handsontable td.read-only-cell {
  background-color: #f8fafc !important;
  color: var(--color-text-secondary);
}

/* Total */
.ht-spreadsheet .handsontable td.total-row-cell {
  background-color: #f0fdfa !important;
  font-weight: 700;
  color: var(--color-primary);
}

/* Categoria */
.ht-spreadsheet .handsontable td.category-cell {
  background-color: #f8fafc !important;
  font-weight: 600;
  padding-left: 12px !important;
}

/* Negativo */
.ht-spreadsheet .handsontable td.negative-value {
  color: #dc2626 !important;
  font-weight: 600;
}

/* Zero */
.ht-spreadsheet .handsontable td.zero-value {
  color: #94a3b8 !important;
}

/* Acima do teto */
.ht-spreadsheet .handsontable td.over-budget {
  background-color: #fef9c3 !important;
}

/* Indicador de formula */
.ht-spreadsheet .handsontable td.formula-cell::after {
  content: '';
  position: absolute;
  top: 1px; right: 1px;
  width: 0; height: 0;
  border-left: 6px solid transparent;
  border-top: 6px solid var(--color-primary);
}

/* Highlight de referencia */
.ref-highlight {
  outline: 2px solid #3b82f6 !important;
  outline-offset: -2px;
  background-color: rgba(59, 130, 246, 0.08) !important;
}

/* Autocomplete */
.formula-autocomplete {
  position: absolute; z-index: 9999;
  background: var(--color-surface);
  border: 1px solid var(--color-border);
  border-radius: 8px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.12);
  max-height: 220px; overflow-y: auto;
  min-width: 260px;
}
.formula-autocomplete-item {
  padding: 6px 12px; cursor: pointer;
  display: flex; flex-direction: column; gap: 1px;
}
.formula-autocomplete-item:hover,
.formula-autocomplete-item.active {
  background-color: var(--color-primary-bg);
}
```

---

## 12. afterChange — Persistencia e Auditoria

O hook `afterChange` do Handsontable e o ponto central para persistir dados e registrar auditoria. Toda alteracao de celula deve ser capturada aqui.

**Regras obrigatorias:**
- Ignorar `source === 'loadData'` (carga inicial, nao e edicao do usuario)
- Ignorar celulas protegidas (total, categorias, etc.)
- Para formulas, persistir o **valor calculado** (`hot.getDataAtCell`), nao a formula em si
- Chamar `hideAutocomplete()` e `clearNavigation()` para limpar estado de edicao
- **Registrar log de auditoria** com: data/hora da alteracao, usuario que fez, celula alterada, valor anterior e valor novo

```tsx
const handleAfterChange = useCallback((
  changes: Handsontable.CellChange[] | null, source: string
) => {
  if (!changes || source === 'loadData') return
  hideAutocomplete()
  clearNavigation()
  const hot = hotRef.current?.hotInstance
  if (!hot) return

  for (const [row, col, oldVal, newVal] of changes) {
    if (oldVal === newVal) continue
    // Pular celulas protegidas
    // ...

    const nv = String(newVal ?? '')
    const finalVal = nv.startsWith('=') ? hot.getDataAtCell(row, col) : newVal

    // Persistir no backend + registrar auditoria
    // Auditoria DEVE conter: timestamp, usuario, celula, oldVal, newVal
  }
}, [])
```

---

## 13. Cleanup de Listeners

```tsx
const handleAfterDeselect = useCallback(() => {
  const textarea = cellEditorInputRef.current
  if (textarea) {
    if (cellEditorListenerRef.current)
      textarea.removeEventListener('input', cellEditorListenerRef.current)
    if (cellEditorKeyListenerRef.current)
      textarea.removeEventListener('keydown',
        cellEditorKeyListenerRef.current as EventListener)
  }
  cellEditorInputRef.current = null
  hideAutocomplete()
  clearNavigation()
  isFormulaMode.current = false
  isEditingInline.current = false
}, [hideAutocomplete, clearNavigation])
```

---

## 14. Re-render ao Mudar Highlight

```tsx
useEffect(() => {
  const hot = hotRef.current?.hotInstance
  if (hot) hot.render()
}, [highlightedRange])
```

---

## 15. Erros Comuns e Solucoes

| Erro | Causa | Solucao |
|------|-------|---------|
| `#NAME?` em formulas localizadas | Handsontable sobrescreve idioma para enGB | Passar CLASSE HyperFormula com `language: 'ptPT'` dentro de `engine` |
| Tela em branco | `height="100%"` sem container de altura fixa | Usar `height="auto"` |
| "Language not registered" | `registerLanguage()` nao chamado no top-level | Chamar fora do componente |
| Setas inserem ref quando deviam mover cursor | `isAtRefInsertionPoint` muito permissivo | Verificar charBefore E charAfter no meio do texto |
| Referencia apaga o resto da formula | `updateNavRef` nao preserva sufixo | Capturar `refSuffix` ao iniciar navegacao |
| Autocomplete nao funciona no editor inline | Listeners nao adicionados ao textarea do HT | Capturar `(editor as any).TEXTAREA` no `afterBeginEditing` |
| Stale closures no autocomplete | useState capturado no closure do listener | Usar refs (`acItemsRef`, `acIdxRef`) |
| Fill handle so vertical | `fillHandle={{ direction: 'vertical' }}` | Usar `fillHandle={true}` |
| Formulas deslocadas | Header como row 0 nos dados | Usar `colHeaders` do HT, dados comecam na row 0 |
| `=SOM` nao mostra `SOMASE` | `FORMULA_DICT` incompleto | Incluir TODAS as variantes condicionais no dicionario (ver secao 8) |
| `=SOMA(` + seta nao insere ref | `getCursorPos` usa `document.activeElement` que falha no `beforeKeyDown` do HT | Checar `isEditingInline + cellEditorInputRef` PRIMEIRO, sem `activeElement` (ver secao 9.3) |

---

## 16. Undo/Redo Confiavel (incluindo formulas)

### Problema

O Undo nativo do Handsontable pode nao refletir corretamente fluxos customizados (formula bar, interceptacao de teclado, persistencia assincrona), especialmente quando a ultima edicao foi formula.

### Solucao

Manter listener global de Ctrl/Cmd+Z e Ctrl/Cmd+Y em capture phase:
- Se editor inline estiver aberto, deixar o browser/editor tratar
- Fora da edicao inline, usar stack custom (old/new source por celula) como prioridade, com fallback no plugin nativo `undoRedo`

### Padrao tecnico

1. Capturar snapshot em `beforeChange`
2. Consolidar entrada no historico em `afterChange` (somente colunas editaveis, sem totals, sem source `loadData`/`customUndo`/`customRedo`)
3. Em undo/redo custom usar `setDataAtCell(..., 'customUndo'|'customRedo')` para evitar loops de historico
4. **NUNCA** fazer update de estado que recarregue a grade inteira a cada tecla (isso destroi o stack de undo)

---

## 17. Persistencia de Formulas no DB (nao so runtime)

### Problema

`afterChange` pode trazer valor computado, e nao necessariamente a formula digitada. Se salvar so numero na tabela principal, ao F5 a formula some.

### Solucao

Criar tabela especifica de formulas por celula:

```sql
-- Exemplo: tabela de formulas para planilha de comissoes
CREATE TABLE CELULA_FORMULAS (
  ID INT AUTO_INCREMENT PRIMARY KEY,
  VENDEDOR_ID INT,
  COLUNA VARCHAR(50),
  FORMULA VARCHAR(500),
  UNIQUE (VENDEDOR_ID, COLUNA)
)
```

### Fluxo correto

**No save:**
1. Detectar formula por `getSourceDataAtCell(row, col)` (fallback: `newVal` string iniciando com `=`)
2. Se formula: upsert em `CELULA_FORMULAS`
3. Se valor numerico: salvar numero na tabela de dados + remover formula antiga dessa celula

**No load:**
1. Buscar dados base + formulas em paralelo
2. Ao montar data da planilha, usar formula salva quando existir; senao valor numerico

### Regra importante

Persistencia assincrona nao pode resetar estado estrutural da grade durante edicao/undo.

---

## 18. "1 clique + = + seta" deve entrar em edicao de formula

### Problema

Com 1 clique, o usuario ainda esta em modo selecao. Ao digitar `=`, dependendo do timing, o editor nao entra de fato; a seta e tratada como navegacao da grade e nao como referencia de formula.

### Solucao (Handsontable way)

No `beforeKeyDown`, ao detectar `=`:

```tsx
// Interceptar = para forcar modo edicao
hot.addHook('beforeKeyDown', (e: KeyboardEvent) => {
  if (e.key === '=' && !hot.getActiveEditor()?.isOpened()) {
    e.preventDefault()
    e.stopPropagation()
    e.stopImmediatePropagation()
    const editor = hot.getActiveEditor()
    editor?.beginEditing('=')
    // Garantir textarea com = e cursor na posicao 1
    const textarea = (editor as any)?.TEXTAREA
    if (textarea) {
      textarea.value = '='
      textarea.setSelectionRange(1, 1)
      // CRITICO: salvar ref do textarea imediatamente — afterBeginEditing
      // pode nao ter rodado quando a proxima tecla chegar
      cellEditorInputRef.current = textarea
    }
    isFormulaMode.current = true
    isEditingInline.current = true
  }
})
```

**CRITICO:** Salvar `cellEditorInputRef.current` e setar `isFormulaMode` + `isEditingInline` dentro do proprio interceptor de `=`. Nao esperar o `afterBeginEditing` — quando a Arrow chega logo em seguida, o hook pode nao ter rodado ainda, e `cellEditorInputRef.current` ainda seria `null`, fazendo `getFormulaText()` e `getCursorPos()` lerem do fallback errado.

### Fallback no beforeKeyDown: detectar formula mode pelo DOM

O `isFormulaMode` pode estar `false` mesmo com o editor aberto e com `=` (ex: beginEditing programatico onde o `onInput` nao rodou). No inicio do handler de `beforeKeyDown` para Arrow/navegacao, adicionar fallback:

```tsx
// Se isFormulaMode nao esta ativo mas o editor tem "=" (beginEditing programatico),
// forcar o modo — nao depender so do onInput que pode nao ter rodado
if (!isFormulaMode.current) {
  const editor = hot.getActiveEditor()
  const textarea = (editor as any)?.TEXTAREA as HTMLTextAreaElement | undefined
  if (textarea && textarea.value.startsWith('=') && (editor as any)?.isOpened?.()) {
    isFormulaMode.current = true
    isEditingInline.current = true
    cellEditorInputRef.current = textarea
  } else {
    return // nao esta em formula mode, deixar HT tratar
  }
}
```

### Guard de editorOpen com fallback para isEditingInline

`isOpened()` pode retornar `false` logo apos `beginEditing()` programatico. Adicionar guard extra:

```tsx
const editorOpen = (hot.getActiveEditor() as any)?.isOpened?.()
const inlineEditing = isEditingInline.current && cellEditorInputRef.current != null
if (!editorOpen && !formulaBarFocused && !inlineEditing) return
```

Sem isso, mesmo com `isFormulaMode = true`, o handler de Arrow retorna cedo porque acha que nenhum editor esta aberto.

### Adicionar fallback no `keydown` do proprio textarea do editor para `Arrow*`
- Se conteudo comeca com `=` e nao esta em text-mode forcado: bloquear propagacao, atualizar referencia (A1, B3:C8 etc.), manter foco no editor

### Regra de ouro

Nao depender so de `input` event para detectar "formula mode"; em alguns fluxos o valor pode ser inserido programaticamente. Sempre ter fallback por leitura direta do `textarea.value`.

### Resumo dos 3 detalhes de timing criticos

| Detalhe | Por que e necessario | Onde aplicar |
|---------|---------------------|--------------|
| Salvar `cellEditorInputRef` no interceptor de `=` | `afterBeginEditing` pode nao ter rodado quando a proxima tecla chega | Bloco do interceptor de `=` no `beforeKeyDown` |
| Fallback lendo `textarea.value` do DOM | `isFormulaMode` pode estar `false` apos `beginEditing` programatico | Inicio do handler de Arrow no `beforeKeyDown` |
| Guard extra com `isEditingInline` | `isOpened()` pode retornar `false` logo apos `beginEditing()` | Check de "editor aberto" antes de processar Arrow |

---

## 19. Resumo da Arquitetura de Eventos

```
Digitou "=" na celula     -> isFormulaMode = true
  |
  +-- Digitou letras       -> showAutocomplete (se match)
  |   +-- ArrowUp/Down     -> navegar dropdown
  |   +-- Tab              -> inserir funcao
  |
  +-- Arrow key            -> isAtRefInsertionPoint?
  |   +-- SIM              -> updateNavRef (navegar, inserir ref)
  |   |   +-- Shift+Arrow  -> extend range (ancora fixa)
  |   +-- NAO              -> mover cursor de texto
  |
  +-- Click em celula      -> isAtRefInsertionPoint?
  |   +-- SIM              -> updateNavRef (inserir ref)
  |   |   +-- Shift+Click  -> extend range
  |   |   +-- Ctrl+Click   -> ";" + nova ref
  |   |   +-- Drag         -> selecionar range
  |   +-- NAO              -> confirmar edicao
  |
  +-- F2                   -> toggle forceTextMode
  +-- Enter                -> confirmar formula
  +-- Escape               -> cancelar edicao
  +-- Digitar char         -> clearNavigation (proximo arrow reinicia)
```
