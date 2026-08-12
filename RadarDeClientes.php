<?php
/**
 * RADAR DE CLIENTES
 * Painel comercial que  aponta clientes em queda, inativos ou em risco de abandono.
 * Telas (?pagina=): alvo | sem | risco | comparativo.
 */

ini_set('display_errors', 1);
error_reporting(E_ALL);
ini_set('memory_limit', '1024M');   
set_time_limit(300);


require_once __DIR__ . '/menu.php';


// Monta os escopos conforme as permissões do usuário. Cada escopo = uma carteira
// (conjunto de supervisores). Sem nenhuma permissão, chuta pra home.
$escoposDisponiveis = [];
if (in_array('radarclientes.php', $permissoes, true)) {
    $escoposDisponiveis['atacado'] = ['label' => 'Atacado', 'sups' => [4]];
}
if (in_array('radarclientes_varejo', $permissoes, true)) {
    $escoposDisponiveis['varejo'] = ['label' => 'Varejo', 'sups' => [2, 8]];
}
if (!$escoposDisponiveis) {
    header('Location: home.php');
    exit;
}

// Escopo escolhido: validado contra os permitidos (URL adulterada não passa)
$escopo = $_GET['escopo'] ?? '';
if (!isset($escoposDisponiveis[$escopo])) $escopo = array_key_first($escoposDisponiveis);
$supsEscopo = $escoposDisponiveis[$escopo]['sups'];


define('TC_DADOS_FILE', __DIR__ . '/dados.xlsx');
define('TC_CACHE_DIR',  __DIR__ . '/cache_tendencia');
define('TC_CACHE_VER',  'v4');   // bump ao mudar o parser/estrutura -> ignora cache antigo


/**
 * Lê um .xlsx sem lib externa (só zip + simplexml, padrão no WAMP).
 * O xlsx é um zip: as strings ficam num pool (sharedStrings) e a planilha só
 * guarda o índice. Retorna as linhas como [CABECALHO => valor]; 1ª linha = cabeçalho.
 */
function tc_ler_xlsx(string $arquivo): array {
    if (!is_file($arquivo)) return [];
    $zip = new ZipArchive();
    if ($zip->open($arquivo) !== true) return [];

    // 1) Pool de textos (strings ficam separadas, referenciadas por índice)
    $shared = [];
    if (($xml = $zip->getFromName('xl/sharedStrings.xml')) !== false) {
        $sst = @simplexml_load_string($xml);
        if ($sst !== false) {
            foreach ($sst->si as $si) {
                if (isset($si->t)) {                 // <si><t>texto</t></si>
                    $shared[] = (string)$si->t;
                } else {                             // <si><r><t>..</t></r>...</si>
                    $buf = '';
                    foreach ($si->r as $r) $buf .= (string)$r->t;
                    $shared[] = $buf;
                }
            }
        }
    }

    // 2) 1ª planilha (normalmente xl/worksheets/sheet1.xml)
    $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    if ($sheetXml === false) {
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $nome = $zip->getNameIndex($i);
            if (strpos($nome, 'xl/worksheets/') === 0 && substr($nome, -4) === '.xml') {
                $sheetXml = $zip->getFromName($nome);
                break;
            }
        }
    }
    $zip->close();
    if (!$sheetXml) return [];

    $sheet = @simplexml_load_string($sheetXml);
    if ($sheet === false) return [];

    // 3) Percorre linhas/células
    $linhas = [];
    $cabec  = null;
    foreach ($sheet->sheetData->row as $row) {
        $celulas = [];
        foreach ($row->c as $c) {
            $ref = (string)$c['r'];                              // ex.: "B12"
            $col = tc_col_index(preg_replace('/\d+/', '', $ref)); // "B" -> 1
            $t   = (string)$c['t'];
            if ($t === 's') {                                    // índice no pool
                $val = $shared[(int)$c->v] ?? '';
            } elseif ($t === 'inlineStr') {
                $val = (string)$c->is->t;
            } else {                                             // número / texto direto
                $val = isset($c->v) ? (string)$c->v : '';
            }
            $celulas[$col] = $val;
        }
        if ($cabec === null) {                                   // 1ª linha = cabeçalho
            $cabec = [];
            foreach ($celulas as $i => $v) $cabec[$i] = trim($v);
            continue;
        }
        $assoc = [];
        foreach ($cabec as $i => $nome) $assoc[$nome] = $celulas[$i] ?? '';
        $linhas[] = $assoc;
    }
    return $linhas;
}

/** Letra(s) de coluna do Excel -> índice 0-based ("A"->0, "B"->1, "AA"->26). */
function tc_col_index(string $letras): int {
    $n = 0;
    $letras = strtoupper($letras);
    for ($i = 0, $len = strlen($letras); $i < $len; $i++) {
        $n = $n * 26 + (ord($letras[$i]) - 64);
    }
    return $n - 1;
}


/**
 * Dataset bruto já normalizado (tipos certos, DATA nula p/ quem nunca comprou).
 * Cacheia em memória (static) e em JSON no disco. O cache é chaveado pelo mtime
 * do arquivo, então trocar o dados.xlsx invalida sozinho. ?atualizar=1 força releitura.
 */
function tc_dados_brutos(): array {
    static $memo = null;
    if ($memo !== null) return $memo;

    $arquivo = TC_DADOS_FILE;
    $mtime   = is_file($arquivo) ? filemtime($arquivo) : 0;
    $cacheF  = TC_CACHE_DIR . '/parse_' . TC_CACHE_VER . '_' . $mtime . '.json';

    if (empty($_GET['atualizar']) && is_file($cacheF)) {
        $j = json_decode((string)file_get_contents($cacheF), true);
        if (is_array($j)) return $memo = $j;
    }

    $rows = [];
    foreach (tc_ler_xlsx($arquivo) as $r) {
        $data = trim((string)($r['DATA_STR'] ?? ''));
        $dias = trim((string)($r['DIAS_ULTIMA_COMPRA'] ?? ''));
        $p80  = trim((string)($r['P80_DIAS_COMPRA'] ?? ''));
        $rows[] = [
            'DATA'               => $data !== '' ? $data : null,     // YYYY-MM-DD ou null
            'COD_CLI'            => (int)($r['COD_CLI'] ?? 0),
            'NOME_CLIENTE'       => (string)($r['NOME_CLIENTE'] ?? ''),
            'CIDADE'             => (string)($r['CIDADE'] ?? ''),
            'COD_RCA'            => (int)($r['COD_RCA'] ?? 0),
            'NOME_RCA'           => (string)($r['NOME_RCA'] ?? 'Sem Vendedor'),
            'COD_SUPERVISOR'     => (int)($r['COD_SUPERVISOR'] ?? 0),
            'NOME_SUPERVISOR'    => (string)($r['NOME_SUPERVISOR'] ?? 'Sem Supervisor'),
            'VALOR_VENDA'        => (float)($r['VALOR_VENDA'] ?? 0),
            'DIAS_ULTIMA_COMPRA' => $dias !== '' ? (int)$dias : null,
            'P80_DIAS_COMPRA'    => $p80  !== '' ? (int)$p80  : null,
            'NUNCA_COMPROU'      => (int)($r['NUNCA_COMPROU'] ?? 0) === 1,
        ];
    }

    if ($rows) {
        if (!is_dir(TC_CACHE_DIR)) @mkdir(TC_CACHE_DIR, 0775, true);
        @file_put_contents($cacheF, json_encode($rows));
    }
    return $memo = $rows;
}




/** Dataset das telas alvo/sem/risco: 24 meses de histórico, filtrado pelo escopo. */
function tc_load_principal(string $dt_fim_mais1_ddmmyyyy, array $sups): array {
    $sups = array_map('intval', $sups);


    // Janela: 24 meses a contar do 1º dia do mês atual, até a data-fim escolhida.
    // Clientes sem pedido entram sempre (não têm data pra filtrar).
    $ini   = (new DateTime('first day of this month'))->modify('-24 months')->format('Y-m-d');
    $fimDt = DateTime::createFromFormat('d/m/Y', $dt_fim_mais1_ddmmyyyy);
    $fim   = $fimDt ? $fimDt->format('Y-m-d') : (new DateTime('tomorrow'))->format('Y-m-d');

    $rows = [];
    foreach (tc_dados_brutos() as $r) {
        if (!in_array($r['COD_SUPERVISOR'], $sups, true)) continue;
        if ($r['DATA'] === null) { $rows[] = $r; continue; }        // sem pedido: sempre entra
        if ($r['DATA'] >= $ini && $r['DATA'] < $fim) $rows[] = $r;  // datas ISO -> comparação textual
    }
    return [$rows, null];
}


/** Dataset da tela comparativo: 3 anos civis (ano_ref-2 até hoje), pelo escopo. */
function tc_load_comparativo(int $ano_ref, array $sups): array {
    $sups = array_map('intval', $sups);
    $ini  = sprintf('%04d-01-01', $ano_ref - 2);
    $fim  = (new DateTime('tomorrow'))->format('Y-m-d');

    $rows = [];
    foreach (tc_dados_brutos() as $r) {
        if (!in_array($r['COD_SUPERVISOR'], $sups, true)) continue;
        if ($r['DATA'] === null) { $rows[] = $r; continue; }        // sem pedido: sempre entra
        if ($r['DATA'] >= $ini && $r['DATA'] < $fim) {
            // o comparativo não usa dias/P80 — zera para manter o mesmo shape da versão antiga
            $r['DIAS_ULTIMA_COMPRA'] = null;
            $r['P80_DIAS_COMPRA']    = null;
            $rows[] = $r;
        }
    }
    return [$rows, null];
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────
const TC_MESES_ABREV = [1=>'Jan',2=>'Fev',3=>'Mar',4=>'Abr',5=>'Mai',6=>'Jun',7=>'Jul',8=>'Ago',9=>'Set',10=>'Out',11=>'Nov',12=>'Dez'];
const TC_MESES_NOME  = [1=>'Janeiro',2=>'Fevereiro',3=>'Março',4=>'Abril',5=>'Maio',6=>'Junho',7=>'Julho',8=>'Agosto',9=>'Setembro',10=>'Outubro',11=>'Novembro',12=>'Dezembro'];

// Formatação BR (milhar com ponto, decimal com vírgula).
function tc_fmt_num($n): string  { return number_format((float)$n, 0, ',', '.'); }
function tc_fmt_brl($n): string  { return 'R$ ' . number_format((float)$n, 0, ',', '.'); }
function tc_fmt_brl2($n): string { return 'R$ ' . number_format((float)$n, 2, ',', '.'); }

// Filtra as linhas pelos vendedores (RCA) selecionados; vazio = todos.
function tc_apply_rca(array $rows, array $rcas): array {
    if (!$rcas) return $rows;
    return array_values(array_filter($rows, fn($r) => in_array($r['NOME_RCA'], $rcas, true)));
}

function tc_rca_options(array $rows): array {
    $s = [];
    foreach ($rows as $r) if ($r['NOME_RCA'] !== '') $s[$r['NOME_RCA']] = true;
    $out = array_keys($s);
    sort($out, SORT_STRING | SORT_FLAG_CASE);
    return $out;
}

/** Lista de meses (ano, mês) do calendário entre duas datas. */
function tc_meses_periodo(DateTime $ini, DateTime $fim): array {
    $out = [];
    $cur = new DateTime($ini->format('Y-m-01'));
    $end = new DateTime($fim->format('Y-m-01'));
    while ($cur <= $end) {
        $out[] = [(int)$cur->format('Y'), (int)$cur->format('n')];
        $cur->modify('first day of next month');
    }
    return $out;
}

// ─────────────────────────────────────────────────────────────────────────────
// PROCESSAMENTO — port fiel do process_queda (app.py)
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Núcleo da análise de queda. Pivota faturamento por cliente × período, compara o
 * período atual contra a referência (mês anterior, média dos meses, ou trimestre)
 * e devolve só quem caiu (var <= 0). Com 6 meses selecionados, agrupa em 2 trimestres.
 */
function tc_process_queda(array $rows, DateTime $dtIni, DateTime $dtFim): array {
    // Info por cliente: DIAS_ULTIMA_COMPRA (mín) e P80 (primeiro não nulo)
    $info = [];
    $zerados = [];   // clientes que nunca tiveram pedido (DATA nula)
    foreach ($rows as $r) {
        $c = $r['COD_CLI'];
        if (!isset($info[$c])) $info[$c] = ['dias' => null, 'p80' => null];
        if ($r['DIAS_ULTIMA_COMPRA'] !== null && ($info[$c]['dias'] === null || $r['DIAS_ULTIMA_COMPRA'] < $info[$c]['dias']))
            $info[$c]['dias'] = $r['DIAS_ULTIMA_COMPRA'];
        if ($info[$c]['p80'] === null && $r['P80_DIAS_COMPRA'] !== null)
            $info[$c]['p80'] = $r['P80_DIAS_COMPRA'];
        if ($r['DATA'] === null && !isset($zerados[$c])) $zerados[$c] = $r;
    }

    $periodos = tc_meses_periodo($dtIni, $dtFim);
    if (count($periodos) < 2) return [[], [], null, '', false, 'Selecione um intervalo com pelo menos 2 meses para análise de tendência.'];

    $modoTri = (count($periodos) === 6);

    if ($modoTri) {
        $b1 = array_slice($periodos, 0, 3);
        $b2 = array_slice($periodos, 3);
        $lbl = function (array $b): string {
            [$a0, $m0] = $b[0]; [$a1, $m1] = $b[count($b) - 1];
            return $a0 === $a1
                ? TC_MESES_ABREV[$m0] . '-' . TC_MESES_ABREV[$m1] . '/' . $a1
                : TC_MESES_ABREV[$m0] . '/' . $a0 . '-' . TC_MESES_ABREV[$m1] . '/' . $a1;
        };
        $l1 = $lbl($b1); $l2 = $lbl($b2);
        $mapa = [];
        foreach ($b1 as [$a, $m]) $mapa["$a-$m"] = $l1;
        foreach ($b2 as [$a, $m]) $mapa["$a-$m"] = $l2;
        $ordem = [$l1, $l2];
    } else {
        $mapa = []; $ordem = [];
        foreach ($periodos as [$a, $m]) {
            $l = TC_MESES_ABREV[$m] . '/' . $a;
            $mapa["$a-$m"] = $l;
            $ordem[] = $l;
        }
    }

    // Pivot: soma de VALOR_VENDA por cliente × período
    $piv = [];
    $iniStr = $dtIni->format('Y-m-d');
    $fimStr = $dtFim->format('Y-m-d');
    foreach ($rows as $r) {
        if ($r['DATA'] === null) continue;
        if ($r['DATA'] < $iniStr || $r['DATA'] > $fimStr) continue;
        $a = (int)substr($r['DATA'], 0, 4);
        $m = (int)substr($r['DATA'], 5, 2);
        $l = $mapa["$a-$m"] ?? null;
        if ($l === null) continue;
        $c = $r['COD_CLI'];
        if (!isset($piv[$c])) {
            $piv[$c] = [
                'COD_CLI' => $c, 'NOME_CLIENTE' => $r['NOME_CLIENTE'], 'CIDADE' => $r['CIDADE'],
                'COD_RCA' => $r['COD_RCA'], 'NOME_RCA' => $r['NOME_RCA'],
                'NOME_SUPERVISOR' => $r['NOME_SUPERVISOR'], 'COD_SUPERVISOR' => $r['COD_SUPERVISOR'],
                'vals' => array_fill_keys($ordem, 0.0),
            ];
        }
        $piv[$c]['vals'][$l] += $r['VALOR_VENDA'];
    }

    $mesAtual  = $ordem[count($ordem) - 1];
    $mesesHist = array_slice($ordem, 0, -1);
    $labelMedia = count($mesesHist) > 1 ? 'Média' : null;

    $resultado = [];
    foreach ($piv as $c => $p) {
        $vHist = array_map(fn($l) => $p['vals'][$l], $mesesHist);
        $ref   = count($vHist) ? array_sum($vHist) / count($vHist) : 0.0;
        $atual = $p['vals'][$mesAtual];
        $var   = $ref == 0.0 ? 100.0 : round(($atual - $ref) / $ref * 100, 1);
        if ($var > 0) continue;               // resultado = resultado[Variação <= 0]

        $linha = $p;
        $linha['media']         = $labelMedia ? round($ref, 2) : null;
        $linha['atual']         = $atual;
        $linha['var']           = $var;
        $linha['p80']           = $info[$c]['p80']  ?? null;
        $linha['dias']          = $info[$c]['dias'] ?? null;
        $linha['NUNCA_COMPROU'] = false;
        $resultado[] = $linha;
    }

    // Reincorpora clientes que nunca compraram: tudo zerado, var = 0
    foreach ($zerados as $c => $r) {
        $resultado[] = [
            'COD_CLI' => $c, 'NOME_CLIENTE' => $r['NOME_CLIENTE'], 'CIDADE' => $r['CIDADE'],
            'COD_RCA' => $r['COD_RCA'], 'NOME_RCA' => $r['NOME_RCA'],
            'NOME_SUPERVISOR' => $r['NOME_SUPERVISOR'], 'COD_SUPERVISOR' => $r['COD_SUPERVISOR'],
            'vals' => array_fill_keys($ordem, 0.0),
            'media' => $labelMedia ? 0.0 : null, 'atual' => 0.0, 'var' => 0.0,
            'p80' => null, 'dias' => null, 'NUNCA_COMPROU' => true,
        ];
    }

    return [$resultado, $mesesHist, $labelMedia, $mesAtual, $modoTri, null];
}


/** Igual ao process_queda, mas para períodos livres (meses/tri/semestres) ano a ano. */
function tc_process_comparativo(array $rows, array $mapa, array $ordem): array {
    $zerados = [];
    $piv = [];
    foreach ($rows as $r) {
        $c = $r['COD_CLI'];
        if (!empty($r['NUNCA_COMPROU'])) { if (!isset($zerados[$c])) $zerados[$c] = $r; continue; }
        if ($r['DATA'] === null) continue;
        $a = (int)substr($r['DATA'], 0, 4);
        $m = (int)substr($r['DATA'], 5, 2);
        $l = $mapa["$a-$m"] ?? null;
        if ($l === null) continue;
        if (!isset($piv[$c])) {
            $piv[$c] = [
                'COD_CLI' => $c, 'NOME_CLIENTE' => $r['NOME_CLIENTE'], 'CIDADE' => $r['CIDADE'],
                'COD_RCA' => $r['COD_RCA'], 'NOME_RCA' => $r['NOME_RCA'],
                'NOME_SUPERVISOR' => $r['NOME_SUPERVISOR'], 'COD_SUPERVISOR' => $r['COD_SUPERVISOR'],
                'vals' => array_fill_keys($ordem, 0.0),
            ];
        }
        $piv[$c]['vals'][$l] += $r['VALOR_VENDA'];
    }

    $periodoAtual = $ordem[count($ordem) - 1];
    $periodosHist = array_slice($ordem, 0, -1);
    $labelMedia   = count($periodosHist) > 1 ? 'Média' : null;

    $resultado = [];
    foreach ($piv as $p) {
        if (array_sum($p['vals']) <= 0) continue;   // movimento em ao menos um período
        // Referência = apenas o ano anterior (último período histórico); o ano-2 fica só como informação no gráfico/tabela
        $ref   = $periodosHist ? $p['vals'][$periodosHist[count($periodosHist) - 1]] : 0.0;
        $atual = $p['vals'][$periodoAtual];
        $p['media']         = $labelMedia ? round($ref, 2) : null;
        $p['atual']         = $atual;
        $p['var']           = $ref == 0.0 ? 100.0 : round(($atual - $ref) / $ref * 100, 1);
        $p['NUNCA_COMPROU'] = false;
        $resultado[] = $p;
    }

    foreach ($zerados as $c => $r) {
        $resultado[] = [
            'COD_CLI' => $c, 'NOME_CLIENTE' => $r['NOME_CLIENTE'], 'CIDADE' => $r['CIDADE'],
            'COD_RCA' => $r['COD_RCA'], 'NOME_RCA' => $r['NOME_RCA'],
            'NOME_SUPERVISOR' => $r['NOME_SUPERVISOR'], 'COD_SUPERVISOR' => $r['COD_SUPERVISOR'],
            'vals' => array_fill_keys($ordem, 0.0),
            'media' => $labelMedia ? 0.0 : null, 'atual' => 0.0, 'var' => 0.0, 'NUNCA_COMPROU' => true,
        ];
    }

    return [$resultado, $periodosHist, $labelMedia, $periodoAtual];
}

// ─────────────────────────────────────────────────────────────────────────────
// PARÂMETROS DA TELA
// ─────────────────────────────────────────────────────────────────────────────
$hoje = new DateTime('today');

// Janela padrão da tela: últimos 3 meses fechados (até o fim do mês passado).
$primeiroDiaMesAtual = new DateTime($hoje->format('Y-m-01'));
$ultimoDiaMesPassado = (clone $primeiroDiaMesAtual)->modify('-1 day');
$defIni = (clone $ultimoDiaMesPassado)->modify('-2 months')->modify('first day of this month');

// Aba ativa, validada contra a lista (URL adulterada cai no default 'alvo').
$paginasValidas = ['alvo', 'sem', 'risco', 'comparativo'];
$pagina = in_array($_GET['pagina'] ?? '', $paginasValidas, true) ? $_GET['pagina'] : 'alvo';

$dtIni = DateTime::createFromFormat('Y-m-d', $_GET['dt_ini'] ?? '') ?: clone $defIni;
$dtFim = DateTime::createFromFormat('Y-m-d', $_GET['dt_fim'] ?? '') ?: clone $ultimoDiaMesPassado;
$dtIni->setTime(0, 0); $dtFim->setTime(0, 0);
if ($dtFim > $hoje) $dtFim = clone $hoje;

// Página de risco usa período fixo de 365 dias
if ($pagina === 'risco') {
    $dtIni = (clone $hoje)->modify('-364 days');
    $dtFim = clone $hoje;
}

$rcasSel = array_values(array_filter((array)($_GET['rca'] ?? []), 'strlen'));

// ── Carrega dados ────────────────────────────────────────────────────────────
$erroConexao = null;
$raw = [];
$rawComp = [];

if ($pagina === 'comparativo') {
    [$rawComp, $erroConexao] = tc_load_comparativo((int)$hoje->format('Y'), $supsEscopo);
    $rcaOpts = tc_rca_options($rawComp);
} else {
    $dtFimMais1 = (clone $dtFim)->modify('+1 day')->format('d/m/Y');
    [$raw, $erroConexao] = tc_load_principal($dtFimMais1, $supsEscopo);
    $rcaOpts = tc_rca_options($raw);
}

$avisos = [];
$dadosPagina = null;   // preenchido por página abaixo

// ─────────────────────────────────────────────────────────────────────────────
// LÓGICA POR PÁGINA (calcula tudo antes do HTML)
// ─────────────────────────────────────────────────────────────────────────────
if (!$erroConexao) {

    if ($pagina === 'alvo') {
        $df = tc_apply_rca($raw, $rcasSel);
        if ($dtIni >= $dtFim) {
            $avisos[] = ['warn', '⚠️ A data inicial deve ser anterior à data final.'];
        } else {
            $ultimoDia = (int)$dtFim->format('t');
            if ((int)$dtFim->format('j') < $ultimoDia) {
                $avisos[] = ['warn', '⚠️ O mês de ' . $dtFim->format('m/Y') . ' está incompleto (' . $dtFim->format('j') . ' dias). Considere selecionar até o último dia do mês anterior para análise mais precisa.'];
            }
            [$res, $mesesHist, $labelMedia, $mesAtual, $modoTri, $erroQ] = tc_process_queda($df, $dtIni, $dtFim);
            if ($erroQ) {
                $avisos[] = ['warn', '⚠️ ' . $erroQ];
            } else {
                // Foco da tela: só quedas relevantes (entre -10% e -100%) + quem nunca comprou.
                // Fica de fora quem caiu pouco (ruído) e quem caiu -100% já sumiu de vez.
                $res = array_values(array_filter($res, fn($r) =>
                    ($r['var'] < -10 && $r['var'] > -100) || $r['NUNCA_COMPROU']
                ));
                if (!$res) {
                    $avisos[] = ['info', 'ℹ️ Nenhum cliente em queda para os filtros selecionados.'];
                } else {
                    usort($res, fn($a, $b) => $a['var'] <=> $b['var']);

                    $somaRef = 0.0; $somaAtual = 0.0; $somaVar = 0.0; $rcasAfetados = [];
                    foreach ($res as $r) {
                        $somaRef   += $labelMedia ? $r['media'] : $r['vals'][$mesesHist[0]];
                        $somaAtual += $r['atual'];
                        $somaVar   += $r['var'];
                        $rcasAfetados[$r['NOME_RCA']] = true;
                    }
                    $perdaTotal = (int)($somaRef - $somaAtual);
                    $varMedia   = count($res) ? $somaVar / count($res) : 0;

                    if ($modoTri) {
                        $txtLogica  = 'Comparação por trimestre: <b>' . $mesesHist[0] . '</b> → <b>' . $mesAtual . '</b>';
                        $labelPerda = 'vs ' . $mesesHist[0];
                        $tituloGraf = 'Comparativo trimestral: ' . $mesesHist[0] . ' vs ' . $mesAtual;
                    } elseif (count($mesesHist) === 1) {
                        $txtLogica  = 'Comparação direta: <b>' . $mesesHist[0] . '</b> → <b>' . $mesAtual . '</b>';
                        $labelPerda = 'vs ' . substr($mesesHist[0], 0, 3);
                        $tituloGraf = $mesesHist[0] . ' vs ' . $mesAtual;
                    } else {
                        $abrevs     = implode(', ', array_map(fn($m) => substr($m, 0, 3), $mesesHist));
                        $txtLogica  = 'Média de <b>' . count($mesesHist) . ' meses</b> (' . $abrevs . ') vs <b>' . $mesAtual . '</b>';
                        $labelPerda = 'vs média histórica';
                        $anoRef     = (int)$dtFim->format('Y');
                        $tituloGraf = 'Média faturamento dos Clientes em queda: ' . ($anoRef - 1) . ' vs ' . $anoRef;
                    }

                    // Gráfico: soma mensal dos clientes em queda — período atual vs mesmo período ano anterior
                    $codsQueda = array_fill_keys(array_column($res, 'COD_CLI'), true);
                    $mesesPer  = tc_meses_periodo($dtIni, $dtFim);
                    $somas = [];   // "ano-mes" => soma
                    foreach ($df as $r) {
                        if ($r['DATA'] === null || !isset($codsQueda[$r['COD_CLI']])) continue;
                        $k = (int)substr($r['DATA'], 0, 4) . '-' . (int)substr($r['DATA'], 5, 2);
                        $somas[$k] = ($somas[$k] ?? 0) + $r['VALOR_VENDA'];
                    }
                    $labelsX = []; $yAtual = []; $yAnterior = [];
                    foreach ($mesesPer as [$a, $m]) {
                        $labelsX[]   = TC_MESES_ABREV[$m];
                        $yAtual[]    = round($somas["$a-$m"] ?? 0, 2);
                        $yAnterior[] = round($somas[($a - 1) . "-$m"] ?? 0, 2);
                    }
                    $anoAtualLbl = $mesesPer[count($mesesPer) - 1][0];
                    if (!array_filter($yAnterior)) {
                        $avisos[] = ['info', '⚠️ Sem dados de ' . ($anoAtualLbl - 1) . ' no arquivo — reexporte o dados.xlsx com mais histórico.'];
                    }

                    $dadosPagina = compact('res', 'mesesHist', 'labelMedia', 'mesAtual', 'modoTri',
                        'perdaTotal', 'varMedia', 'txtLogica', 'labelPerda', 'tituloGraf',
                        'labelsX', 'yAtual', 'yAnterior', 'anoAtualLbl') + ['rcasAfetados' => count($rcasAfetados)];
                }
            }
        }

    } elseif ($pagina === 'sem') {
        $df = tc_apply_rca($raw, $rcasSel);
        if ($dtIni >= $dtFim) {
            $avisos[] = ['warn', '⚠️ A data inicial deve ser anterior à data final.'];
        } else {
            $iniStr = $dtIni->format('Y-m-d'); $fimStr = $dtFim->format('Y-m-d');

            // Última compra por cliente (data máxima; null = nunca comprou)
            $ult = []; $noPeriodo = [];
            foreach ($df as $r) {
                $c = $r['COD_CLI'];
                if ($r['DATA'] !== null && $r['DATA'] >= $iniStr && $r['DATA'] <= $fimStr) $noPeriodo[$c] = true;
                if (!isset($ult[$c]) || ($r['DATA'] !== null && ($ult[$c]['DATA'] === null || $r['DATA'] > $ult[$c]['DATA']))) {
                    if (!isset($ult[$c])) $ult[$c] = $r;
                    elseif ($r['DATA'] !== null) $ult[$c] = $r;
                }
            }

            $semCompra = []; $novosSemPedido = 0; $somaDias = 0; $qtdComHist = 0; $rcasAfetados = [];
            foreach ($ult as $c => $r) {
                if (isset($noPeriodo[$c])) continue;
                $nunca = ($r['DATA'] === null);
                $dias  = $nunca ? 999 : (int)$hoje->diff(new DateTime($r['DATA']))->days;
                if (!$nunca && $dias > 365) continue;     // corte de 365 dias só p/ quem tem histórico
                if ($nunca) $novosSemPedido++;
                else { $somaDias += $dias; $qtdComHist++; }
                $rcasAfetados[$r['NOME_RCA']] = true;
                $semCompra[] = [
                    'COD_CLI' => $c, 'NOME_CLIENTE' => $r['NOME_CLIENTE'], 'CIDADE' => $r['CIDADE'],
                    'COD_RCA' => $r['COD_RCA'], 'NOME_RCA' => $r['NOME_RCA'],
                    'dias' => $nunca ? null : $dias, 'nunca' => $nunca,
                    'ultima' => $r['DATA'], 'p80' => $r['P80_DIAS_COMPRA'],
                ];
            }

            if (!$semCompra) {
                $avisos[] = ['info', 'ℹ️ Nenhum cliente com menos de 365 dias sem comprar no período.'];
            } else {
                // Ordena: maior inatividade primeiro; "nunca comprou" ao final
                usort($semCompra, fn($a, $b) => ($b['dias'] ?? -1) <=> ($a['dias'] ?? -1));
                $totalCarteira = count($ult);
                $pct        = $totalCarteira ? count($semCompra) / $totalCarteira * 100 : 0;
                $mediaDias  = $qtdComHist ? $somaDias / $qtdComHist : 0;
                $diasPeriodo = (int)$dtIni->diff($dtFim)->days + 1;
                $dadosPagina = compact('semCompra', 'pct', 'mediaDias', 'novosSemPedido', 'diasPeriodo')
                             + ['rcasAfetados' => count($rcasAfetados)];
            }
        }

    } elseif ($pagina === 'risco') {
        $df = tc_apply_rca($raw, $rcasSel);

        $agg = [];
        foreach ($df as $r) {           // rows já vêm ordenadas por DATA (nulos primeiro)
            $c = $r['COD_CLI'];
            if (!isset($agg[$c])) {
                $agg[$c] = [
                    'COD_CLI' => $c, 'NOME_CLIENTE' => $r['NOME_CLIENTE'], 'CIDADE' => $r['CIDADE'],
                    'COD_RCA' => $r['COD_RCA'], 'NOME_RCA' => $r['NOME_RCA'],
                    'dias' => $r['DIAS_ULTIMA_COMPRA'], 'p80' => $r['P80_DIAS_COMPRA'], 'total' => 0.0,
                ];
            } else {
                // 'last' do pandas: sobrescreve dados cadastrais com a linha mais recente
                $agg[$c]['NOME_CLIENTE'] = $r['NOME_CLIENTE'];
                $agg[$c]['CIDADE']       = $r['CIDADE'];
                $agg[$c]['COD_RCA']      = $r['COD_RCA'];
                $agg[$c]['NOME_RCA']     = $r['NOME_RCA'];
                if ($r['DIAS_ULTIMA_COMPRA'] !== null && ($agg[$c]['dias'] === null || $r['DIAS_ULTIMA_COMPRA'] < $agg[$c]['dias']))
                    $agg[$c]['dias'] = $r['DIAS_ULTIMA_COMPRA'];
                if ($agg[$c]['p80'] === null && $r['P80_DIAS_COMPRA'] !== null)
                    $agg[$c]['p80'] = $r['P80_DIAS_COMPRA'];
            }
            $agg[$c]['total'] += $r['VALOR_VENDA'];
        }

        $risco = []; $somaP80 = 0; $somaRatio = 0; $somaTotal = 0;
        foreach ($agg as $a) {
            $dias = (int)($a['dias'] ?? 0);
            $p80  = (int)($a['p80'] ?? 0);
            $ratio = $p80 === 0 ? 0.0 : round($dias / $p80, 2);
            // Critério de risco: já passou 30% além do ritmo normal (ratio >= 1.3), mas ainda
            // dá pra reverter (dias < 300), e é cliente que importa (faturou > R$ 5 mil no ano).
            if ($p80 > 0 && $dias > 21 && $dias < 300 && $ratio >= 1.3 && $ratio < 10 && $a['total'] > 5000) {
                $a['dias'] = $dias; $a['p80'] = $p80; $a['ratio'] = $ratio;
                $risco[] = $a;
                $somaP80 += $p80; $somaRatio += $ratio; $somaTotal += $a['total'];
            }
        }

        if (!$risco) {
            $avisos[] = ['info', '✅ Nenhum cliente encontrado com os filtros aplicados.'];
        } else {
            usort($risco, fn($x, $y) => $y['ratio'] <=> $x['ratio']);   // pior ratio primeiro (igual app.py)
            $dadosPagina = [
                'risco'      => $risco,
                'mediaP80'   => $somaP80 / count($risco),
                'mediaRatio' => $somaRatio / count($risco),
                'somaTotal'  => $somaTotal,
            ];
        }

    } else { // comparativo
        $df = tc_apply_rca($rawComp, $rcasSel);

        $anoAtual = (int)$hoje->format('Y');
        $triAtual = intdiv((int)$hoje->format('n') - 1, 3) + 1;
        $modosValidos = ['meses', 'tri', 'tri_ano', 'sem'];
        $modoComp = in_array($_GET['modo'] ?? '', $modosValidos, true) ? $_GET['modo'] : 'meses';

        $mesesSel = array_values(array_filter(array_map('intval', (array)($_GET['meses'] ?? [])), fn($m) => $m >= 1 && $m <= 12));
        if (!$mesesSel && !isset($_GET['modo'])) $mesesSel = [(int)$hoje->format('n')];   // default 1º acesso
        sort($mesesSel);
        $triSel = max(1, min(4, (int)($_GET['tri'] ?? $triAtual)));

        $mapa = []; $ordem = []; $mapaGraf = []; $ordemGraf = []; $txtLogicaC = '';

        // Modo Semestre vs Semestre — anos disponíveis dentro da janela carregada
        $anosOpc = [$anoAtual, $anoAtual - 1, $anoAtual - 2];
        $semBaseSel = $anoBaseSel = $semAtualSel = $anoAtualSel = null;

        if ($modoComp === 'meses') {
            if ($mesesSel) {
                $abrev = implode('+', array_map(fn($m) => TC_MESES_ABREV[$m], $mesesSel));
                $anos  = [$anoAtual - 2, $anoAtual - 1, $anoAtual];
                foreach ($anos as $a) {
                    $l = "$abrev/$a"; $ordem[] = $l;
                    foreach ($mesesSel as $m) $mapa["$a-$m"] = $l;
                }
                $txtLogicaC = "Comparação de <b>$abrev</b> entre <b>{$anos[0]}</b>, <b>{$anos[1]}</b> e <b>{$anos[2]}</b>";
                $mapaGraf = $mapa; $ordemGraf = $ordem;
            } else {
                $avisos[] = ['info', 'ℹ️ Selecione ao menos um mês para continuar.'];
            }
        } elseif ($modoComp === 'tri') {
            $mesesTri = [($triSel - 1) * 3 + 1, ($triSel - 1) * 3 + 2, ($triSel - 1) * 3 + 3];
            $anos = [$anoAtual - 2, $anoAtual - 1, $anoAtual];
            foreach ($anos as $a) {
                $l = "T$triSel/$a"; $ordem[] = $l;
                foreach ($mesesTri as $m) $mapa["$a-$m"] = $l;
            }
            $txtLogicaC = "Comparação de <b>T$triSel</b> entre <b>{$anos[0]}</b>, <b>{$anos[1]}</b> e <b>{$anos[2]}</b>";
            $mapaGraf = $mapa; $ordemGraf = $ordem;
        } elseif ($modoComp === 'tri_ano') { // trimestres do ano atual
            for ($q = 1; $q <= $triAtual; $q++) {
                $l = "T$q/$anoAtual"; $ordemGraf[] = $l;
                foreach ([($q - 1) * 3 + 1, ($q - 1) * 3 + 2, ($q - 1) * 3 + 3] as $m) $mapaGraf["$anoAtual-$m"] = $l;
            }
            $ordem = array_slice($ordemGraf, 0, -1);                     // só trimestres FECHADOS
            $mapa  = array_filter($mapaGraf, fn($l) => in_array($l, $ordem, true));
            if (count($ordem) >= 2) {
                $txtLogicaC = "Trimestres fechados de <b>$anoAtual</b>: " . implode(' → ', array_map(fn($p) => "<b>$p</b>", $ordem));
            } elseif (count($ordem) === 1) {
                $txtLogicaC = "Apenas <b>{$ordem[0]}</b> fechado até agora em $anoAtual — aguarde o próximo trimestre fechar para haver comparação.";
            } else {
                $txtLogicaC = "Ainda não há trimestre fechado em $anoAtual.";
            }
        } else { // sem — Semestre vs Semestre (comparação direta e livre entre dois semestres)
            $semAtual = ((int)$hoje->format('n') <= 6) ? 1 : 2;

            // Seleção validada contra a janela de dados (anos carregados) e semestre 1/2.
            // Defaults: "atual" = semestre corrente; "base" = semestre imediatamente anterior.
            $semBaseSel  = in_array((int)($_GET['sbase']  ?? 0), [1, 2], true)   ? (int)$_GET['sbase']  : ($semAtual === 1 ? 2 : 1);
            $anoBaseSel  = in_array((int)($_GET['abase']  ?? 0), $anosOpc, true) ? (int)$_GET['abase']  : ($semAtual === 1 ? $anoAtual - 1 : $anoAtual);
            $semAtualSel = in_array((int)($_GET['satual'] ?? 0), [1, 2], true)   ? (int)$_GET['satual'] : $semAtual;
            $anoAtualSel = in_array((int)($_GET['aatual'] ?? 0), $anosOpc, true) ? (int)$_GET['aatual'] : $anoAtual;

            $rotulaSem = fn($s, $a) => ($s === 1 ? '1º Sem' : '2º Sem') . "/$a";
            $lBase  = $rotulaSem($semBaseSel,  $anoBaseSel);
            $lAtual = $rotulaSem($semAtualSel, $anoAtualSel);

            if ($semBaseSel === $semAtualSel && $anoBaseSel === $anoAtualSel) {
                $avisos[] = ['info', 'ℹ️ Selecione dois semestres diferentes para comparar.'];
            } else {
                $mesesBase  = $semBaseSel  === 1 ? [1, 2, 3, 4, 5, 6] : [7, 8, 9, 10, 11, 12];
                $mesesAtual = $semAtualSel === 1 ? [1, 2, 3, 4, 5, 6] : [7, 8, 9, 10, 11, 12];
                $ordem = [$lBase, $lAtual];   // base (referência) → atual (comparado)
                foreach ($mesesBase  as $m) $mapa["$anoBaseSel-$m"]  = $lBase;
                foreach ($mesesAtual as $m) $mapa["$anoAtualSel-$m"] = $lAtual;
                $mapaGraf = $mapa; $ordemGraf = $ordem;
                $txtLogicaC = "Comparação direta: <b>$lBase</b> → <b>$lAtual</b>";

                // Aviso se o semestre comparado ainda não fechou (meses em aberto entram parciais)
                $fimAtualSem = (new DateTime(sprintf('%04d-%02d-01', $anoAtualSel, $semAtualSel === 1 ? 6 : 12)))
                               ->modify('last day of this month');
                if ($fimAtualSem >= $hoje) {
                    $avisos[] = ['info', 'ℹ️ O semestre comparado (<b>' . $lAtual . '</b>) ainda não fechou — os meses em aberto entram parciais.'];
                }
            }
        }

        if ($ordemGraf) {
            // Primeiro apura os clientes em queda (os que vão para a tabela);
            // o gráfico de barras abaixo se restringe exatamente a esses clientes.
            $resComp = []; $periodosHistC = []; $labelMediaC = null; $periodoAtualC = '';
            $kpiComp = null;
            if (count($ordem) >= 2) {
                [$resComp, $periodosHistC, $labelMediaC, $periodoAtualC] = tc_process_comparativo($df, $mapa, $ordem);
                $resComp = array_values(array_filter($resComp, fn($r) => $r['var'] < -10 && $r['var'] > -100));
                if ($resComp) {
                    usort($resComp, fn($a, $b) => $a['var'] <=> $b['var']);
                    $somaRefC = 0.0; $somaAtualC = 0.0; $cres = 0; $cai = 0;
                    foreach ($resComp as $r) {
                        $somaRefC   += $labelMediaC ? $r['media'] : $r['vals'][$periodosHistC[0]];
                        $somaAtualC += $r['atual'];
                        if ($r['var'] > 0) $cres++;
                        if ($r['var'] < 0) $cai++;
                    }
                    $kpiComp = [
                        'qtd' => count($resComp), 'cresceram' => $cres, 'cairam' => $cai,
                        'variacao' => $somaRefC ? ($somaAtualC - $somaRefC) / $somaRefC * 100 : 0,
                        'somaRef' => $somaRefC, 'somaAtual' => $somaAtualC,
                    ];
                }
            }

            // Gráfico de barras: soma por período APENAS dos clientes em queda listados na tabela.
            $codsComp = array_fill_keys(array_column($resComp, 'COD_CLI'), true);
            $totais = array_fill_keys($ordemGraf, 0.0);
            foreach ($df as $r) {
                if ($r['DATA'] === null || !isset($codsComp[$r['COD_CLI']])) continue;
                $k = (int)substr($r['DATA'], 0, 4) . '-' . (int)substr($r['DATA'], 5, 2);
                $l = $mapaGraf[$k] ?? null;
                if ($l !== null) $totais[$l] += $r['VALOR_VENDA'];
            }

            $dadosPagina = compact('modoComp', 'mesesSel', 'triSel', 'triAtual', 'anoAtual',
                'semBaseSel', 'anoBaseSel', 'semAtualSel', 'anoAtualSel', 'anosOpc',
                'ordemGraf', 'totais', 'ordem', 'txtLogicaC',
                'resComp', 'periodosHistC', 'labelMediaC', 'periodoAtualC', 'kpiComp');
        }
    }
}

// Querystring base para os links das abas (preserva filtros)
function tc_link_aba(string $pag): string {
    $q = $_GET;
    $q['pagina'] = $pag;
    unset($q['atualizar']);
    return '?' . http_build_query($q);
}
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Radar de Clientes</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<!-- Fontes de ícone usadas pelo menu (renderizarMenu) -->
<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/apexcharts@3.49.1/dist/apexcharts.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js"></script>
<style>
:root {
  --azul-escuro: #022650;
  --azul: #044da2;
  --amarelo: #ffcc00;
  --branco: #f9f9f9;
  --superficie: #ffffff;
  --borda: rgba(4, 77, 162, 0.12);
  --borda-forte: rgba(4, 77, 162, 0.22);
  --texto: #022650;
  --texto-suave: rgba(2, 38, 80, 0.5);
  --sombra: 0 1px 3px rgba(2,38,80,0.07), 0 4px 16px rgba(2,38,80,0.05);
  --sombra-md: 0 2px 8px rgba(2,38,80,0.10), 0 8px 32px rgba(2,38,80,0.07);
  --radius: 10px;
  --radius-sm: 6px;
}

*, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }

body {
  font-family: 'Sora', sans-serif;
  background: linear-gradient(135deg, #f5f7fa 0%, #eef2f6 100%);
  background-attachment: fixed;
  color: var(--texto);
  font-size: 13.5px;
  line-height: 1.5;
  min-height: 100vh;
}

.page-wrap { max-width: 1360px; margin: 0 auto; padding: 96px 32px 32px; }

.page-header { margin-bottom: 20px; }
.page-header h1 { font-size: 28px; font-weight: 700; color: #1f2937; letter-spacing: -0.01em; }
.page-header p  { color: #4b5563; font-size: 14px; margin-top: 2px; }
.badge-escopo {
  display: inline-block;
  vertical-align: middle;
  margin-left: 8px;
  padding: 3px 12px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: .04em;
  text-transform: uppercase;
  background: var(--azul);
  color: #fff;
}

/* ── Barra de abas + filtros (abas à esquerda, filtros à direita) ── */
.toolbar {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 16px;
  flex-wrap: wrap;
  margin-bottom: 18px;
}
.toolbar .tab { margin-bottom: 0; }
/* Painel dos filtros: fundo azulado suave pra diferenciar das abas */
.toolbar form {
  margin-left: auto;                 /* filtros sempre encostados à direita */
  display: flex;
  flex-direction: column;
  align-items: flex-end;
  gap: 8px;
  background: rgba(4, 77, 162, 0.05);
  border: 1px solid var(--borda-forte);
  border-radius: var(--radius);
  padding: 7px 12px 9px;
}
.toolbar .filtros { justify-content: flex-end; gap: 10px; }
/* Controles compactos dentro da barra */
.toolbar .filtro-grupo label { font-size: 9.5px; }
.toolbar .filtro-grupo input[type=date],
.toolbar .filtro-grupo select {
  padding: 5px 7px;
  font-size: 12px;
  min-width: 126px;
}
.toolbar .rca-drop > summary {
  padding: 5px 9px;
  font-size: 12px;
  min-width: 96px;
  max-width: 150px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.toolbar .rca-lista { left: auto; right: 0; }   /* dropdown abre pra dentro da tela */
.toolbar .btn { padding: 6px 13px; font-size: 12px; }

/* ── Filtros do comparativo: tipo à esquerda, demais à direita ── */
.comp-filtros {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 16px;
  flex-wrap: wrap;
  margin-bottom: 18px;
}
/* Painel de destaque (mesmo visual dos filtros da toolbar) */
.painel-destaque {
  margin-left: auto;
  background: rgba(4, 77, 162, 0.05);
  border: 1px solid var(--borda-forte);
  border-radius: var(--radius);
  padding: 7px 12px 9px;
}
.painel-destaque .filtros { justify-content: flex-end; gap: 10px; }
.painel-destaque .filtro-grupo label { font-size: 9.5px; }
.painel-destaque select { padding: 5px 7px; font-size: 12px; min-width: 126px; }
.painel-destaque .meses-check label { padding: 5px 8px; font-size: 11px; }
.painel-destaque .rca-drop > summary {
  padding: 5px 9px;
  font-size: 12px;
  min-width: 96px;
  max-width: 150px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.painel-destaque .rca-lista { left: auto; right: 0; }
.painel-destaque .btn { padding: 6px 13px; font-size: 12px; }

/* Par semestre + ano (modo Semestre vs Semestre) */
.sem-par { display: flex; gap: 5px; }
.painel-destaque .sem-par select { min-width: 0; padding: 5px 4px; font-size: 11px; }
.sem-par select[name="sbase"], .sem-par select[name="satual"] { flex: 0 0 64px; width: 64px; min-width: 64px; }
.sem-par select[name="abase"], .sem-par select[name="aatual"] { flex: 0 0 58px; width: 58px; min-width: 58px; }

/* Altura idêntica para todos os controles do painel de destaque (selects + RCA) */
.painel-destaque select,
.painel-destaque .sem-par select,
.painel-destaque .rca-drop > summary {
  box-sizing: border-box;
  height: 30px;
  line-height: 18px;
}

/* ── Abas (mesmo estilo do agendamentos.php) ── */
.tab { display: flex; gap: 6px; margin-bottom: 18px; flex-wrap: wrap; }
.tab a {
  padding: 9px 16px; border-radius: var(--radius-sm); cursor: pointer;
  font-size: 13px; font-weight: 600; color: var(--texto-suave);
  background: var(--superficie); border: 1px solid var(--borda);
  text-decoration: none; transition: all .15s;
}
.tab a:hover { color: var(--azul); border-color: var(--borda-forte); }
.tab a.active { background: var(--azul); color: #fff; border-color: var(--azul); }

/* ── Cards ── */
.card {
  background: var(--superficie); border: 1px solid var(--borda);
  border-radius: var(--radius); box-shadow: var(--sombra); margin-bottom: 18px;
}
.card-header {
  padding: 14px 18px; border-bottom: 1px solid var(--borda);
  display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap;
}
.card-header h2 { font-size: 15px; font-weight: 600; }
.card-header h2 .obs-graf { font-size: 11px; font-weight: 500; color: var(--azul); margin-left: 8px; }
.card-body { padding: 18px; }

/* ── Filtros ── */
.filtros { display: flex; align-items: flex-end; gap: 14px; flex-wrap: wrap; }
.filtro-grupo { display: flex; flex-direction: column; gap: 4px; }
.filtro-grupo label {
  font-size: 10.5px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .07em; color: var(--texto-suave);
}
.filtro-grupo input[type=date], .filtro-grupo select {
  font-family: 'Sora', sans-serif; font-size: 13px; color: var(--texto);
  padding: 8px 10px; border: 1px solid var(--borda-forte);
  border-radius: var(--radius-sm); background: #fff; min-width: 150px;
}
.btn {
  font-family: 'Sora', sans-serif; font-size: 13px; font-weight: 600;
  padding: 9px 16px; border: 1px solid var(--borda-forte);
  border-radius: var(--radius-sm); background: #fff; color: var(--texto);
  cursor: pointer; transition: all .15s; text-decoration: none; display: inline-block;
}
.btn:hover { border-color: var(--azul); color: var(--azul); }
.btn-primary { background: var(--azul); border-color: var(--azul); color: #fff; }
.btn-primary:hover { background: var(--azul-escuro); color: #fff; }

/* Dropdown de RCA (multiselect) */
.rca-drop { position: relative; }
.rca-drop > summary {
  list-style: none; cursor: pointer; font-size: 13px; padding: 8px 12px;
  border: 1px solid var(--borda-forte); border-radius: var(--radius-sm);
  background: #fff; min-width: 200px; user-select: none;
}
.rca-drop > summary::-webkit-details-marker { display: none; }
.rca-drop[open] > summary { border-color: var(--azul); }
.rca-lista {
  position: absolute; z-index: 30; top: calc(100% + 4px); left: 0;
  background: #fff; border: 1px solid var(--borda-forte); border-radius: var(--radius-sm);
  box-shadow: var(--sombra-md); max-height: 280px; overflow: auto;
  min-width: 260px; padding: 6px;
}
.rca-lista label {
  display: flex; align-items: center; gap: 8px; padding: 6px 8px;
  font-size: 12.5px; border-radius: 4px; cursor: pointer;
  text-transform: none; letter-spacing: 0; font-weight: 500; color: var(--texto);
}
.rca-lista label:hover { background: rgba(4,77,162,0.06); }

/* Checkboxes de meses (comparativo) */
.meses-check { display: flex; gap: 4px; flex-wrap: wrap; }
.meses-check label {
  font-size: 12px; font-weight: 600; padding: 6px 10px; cursor: pointer;
  border: 1px solid var(--borda-forte); border-radius: var(--radius-sm);
  background: #fff; color: var(--texto-suave); user-select: none;
  text-transform: none; letter-spacing: 0;
}
.meses-check input { display: none; }
.meses-check label:has(input:checked) { background: var(--azul); border-color: var(--azul); color: #fff; }

.modo-radio { display: flex; gap: 4px; flex-wrap: wrap; }
.modo-radio label {
  font-size: 12.5px; font-weight: 600; padding: 8px 13px; cursor: pointer;
  border: 1px solid var(--borda-forte); border-radius: var(--radius-sm);
  background: #fff; color: var(--texto-suave); user-select: none;
  text-transform: none; letter-spacing: 0;
}
.modo-radio input { display: none; }
.modo-radio label:has(input:checked) { background: var(--azul); border-color: var(--azul); color: #fff; }

/* ── KPIs (mesmo estilo dos cards .metric do agendamentos.php) ── */
.kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 18px; }
.kpi-card {
  background: var(--superficie);
  border: 1px solid var(--borda);
  border-radius: var(--radius);
  box-shadow: var(--sombra);
  padding: 13px 14px;
  position: relative;
  overflow: hidden;
  transition: box-shadow .2s, transform .2s;
  display: flex;
  align-items: baseline;
  gap: 8px;
  flex-wrap: wrap;
}
.kpi-card:hover { box-shadow: var(--sombra-md); transform: translateY(-1px); }
.kpi-card::before {
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 3px;
  background: var(--azul);
  opacity: 0;
  transition: opacity .2s;
}
.kpi-card:hover::before { opacity: 1; }
.kpi-card:first-child::before { background: var(--amarelo); opacity: 1; }
.kpi-value {
  order: 1;
  font-size: 22px;
  font-weight: 700;
  color: var(--azul-escuro);
  line-height: 1;
  font-variant-numeric: tabular-nums;
}
.kpi-card:first-child .kpi-value { color: var(--azul); }
.kpi-label {
  order: 2;
  font-size: 10.5px;
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: .07em;
  color: var(--texto-suave);
}
.kpi-sub {
  order: 3;
  flex-basis: 100%;
  font-size: 10px;
  color: var(--texto-suave);
  line-height: 1.3;
}

/* ── Avisos ── */
.info-box, .warn-box {
  padding: 12px 16px; border-radius: var(--radius-sm); font-size: 13px; margin-bottom: 14px;
}
.info-box { background: #eef6ff; border: 1px solid #bfdcff; color: #1e4e8c; }
.warn-box { background: #fff8e6; border: 1px solid #ffe1a1; color: #8a6100; }

.section-title { font-size: 15px; font-weight: 700; margin: 20px 0 10px; }

/* ── Tabela ── */
.data-table-wrapper {
  border: 1px solid var(--borda); border-radius: var(--radius);
  overflow: auto; background: #fff; box-shadow: var(--sombra); max-height: 540px;
}
.data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
.data-table th {
  background: #f5f7fa; padding: 11px 12px; text-align: left; font-weight: 700;
  font-size: 10.5px; text-transform: uppercase; letter-spacing: .05em;
  color: var(--azul-escuro); border-bottom: 1px solid var(--borda-forte);
  position: sticky; top: 0; z-index: 5; cursor: pointer; white-space: nowrap; user-select: none;
}
.data-table th .seta { opacity: .45; font-size: 10px; }
.data-table td { padding: 9px 12px; border-bottom: 1px solid var(--borda); white-space: nowrap; }
.td-cliente { max-width: 240px; overflow: hidden; text-overflow: ellipsis; }
.td-cidade  { max-width: 120px; overflow: hidden; text-overflow: ellipsis; }
.td-rca     { max-width: 140px; overflow: hidden; text-overflow: ellipsis; }
.td-codrca  { text-align: center; font-variant-numeric: tabular-nums; }
.data-table th:nth-child(4) { text-align: center; }


#tblSem  th:nth-child(1), #tblSem  td:nth-child(1),
#tblSem  th:nth-child(6), #tblSem  td:nth-child(6)   { text-align: center; }
#tblRisco th:nth-child(n+6), #tblRisco td:nth-child(n+6) { text-align: center; }
#tblComp th:nth-child(1),    #tblComp td:nth-child(1),
#tblComp th:nth-child(n+6),  #tblComp td:nth-child(n+6)  { text-align: center; }
/* Clientes alvo: valores dos meses, Média, mês atual e Var % (6ª em diante) */
#tblAlvo th:nth-child(n+6), #tblAlvo td:nth-child(n+6)   { text-align: center; }
.data-table tbody tr:nth-child(even) { background: #fafbfd; }
.data-table tbody tr:hover { background: rgba(4,77,162,0.05); }
.td-num { text-align: right; font-variant-numeric: tabular-nums; }

.negative-value { background: #FEF2F2 !important; color: #DC2626 !important; font-weight: 600; }
.high-ratio     { background: #FFF7ED !important; color: #EA580C !important; font-weight: 700; }
.medium-ratio   { background: #FFFBEB !important; color: #D97706 !important; font-weight: 600; }
.badge-nunca {
  background: #eef2ff; color: #4338ca; font-size: 10.5px; font-weight: 700;
  padding: 2px 8px; border-radius: 999px;
}

@media (max-width: 900px) {
  .kpi-grid { grid-template-columns: repeat(2, 1fr); }
  .page-wrap { padding: 96px 14px 16px; }
}
</style>
</head>
<body>

<?php

echo renderizarMenu($username, $is_admin, $permissoes, $foto);
?>

<div class="page-wrap">

    <div class="page-header">
        <h1>📡 Radar de Clientes <span class="badge-escopo"><?= $escoposDisponiveis[$escopo]['label'] ?></span></h1>
        <p>Análise de clientes do <?= strtolower($escoposDisponiveis[$escopo]['label']) ?>: queda de compras, inatividade e risco de abandono.</p>
    </div>

    <!-- ── ABAS + FILTROS (mesma linha: abas à esquerda, filtros à direita) ── -->
    <div class="toolbar">
        <div class="tab">
            <a class="<?= $pagina==='alvo'        ? 'active':'' ?>" href="<?= htmlspecialchars(tc_link_aba('alvo')) ?>">📉 Clientes alvo</a>
            <a class="<?= $pagina==='sem'         ? 'active':'' ?>" href="<?= htmlspecialchars(tc_link_aba('sem')) ?>">🚫 Clientes sem compras</a>
            <a class="<?= $pagina==='risco'       ? 'active':'' ?>" href="<?= htmlspecialchars(tc_link_aba('risco')) ?>">📊 Índice de risco</a>
            <a class="<?= $pagina==='comparativo' ? 'active':'' ?>" href="<?= htmlspecialchars(tc_link_aba('comparativo')) ?>">🔄 Comparativo por Período</a>
        </div>

        <?php if ($pagina !== 'comparativo'): ?>
        <form method="get" id="formFiltros">
                <input type="hidden" name="pagina" value="<?= htmlspecialchars($pagina) ?>">
                <div class="filtros">
                    <?php if (count($escoposDisponiveis) > 1): ?>
                    <div class="filtro-grupo">
                        <label>Carteira</label>
                        <select name="escopo" onchange="this.form.submit()">
                            <?php foreach ($escoposDisponiveis as $k => $e): ?>
                            <option value="<?= $k ?>" <?= $k === $escopo ? 'selected' : '' ?>><?= $e['label'] ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <?php else: ?>
                    <input type="hidden" name="escopo" value="<?= htmlspecialchars($escopo) ?>">
                    <?php endif; ?>

                    <?php if ($pagina === 'alvo' || $pagina === 'sem'): ?>
                    <div class="filtro-grupo">
                        <label>Data inicial</label>
                        <input type="date" name="dt_ini" value="<?= $dtIni->format('Y-m-d') ?>" max="<?= $hoje->format('Y-m-d') ?>">
                    </div>
                    <div class="filtro-grupo">
                        <label>Data final</label>
                        <input type="date" name="dt_fim" value="<?= $dtFim->format('Y-m-d') ?>" max="<?= $hoje->format('Y-m-d') ?>">
                    </div>
                    <?php elseif ($pagina === 'risco'): ?>
                    <div class="filtro-grupo">
                        <label>Período</label>
                        <span style="font-size:12px;padding:6px 0;">Últimos 365 dias</span>
                    </div>
                    <?php endif; ?>

                    <div class="filtro-grupo">
                        <label>Vendedor (RCA)</label>
                        <details class="rca-drop">
                            <summary id="rcaResumo">
                                <?= !$rcasSel ? 'Todos' : (count($rcasSel) === 1 ? htmlspecialchars($rcasSel[0]) : count($rcasSel) . ' selecionados') ?>
                            </summary>
                            <div class="rca-lista">
                                <?php foreach ($rcaOpts as $nome): ?>
                                <label><input type="checkbox" name="rca[]" value="<?= htmlspecialchars($nome) ?>" <?= in_array($nome, $rcasSel, true)?'checked':'' ?>> <?= htmlspecialchars($nome) ?></label>
                                <?php endforeach; ?>
                            </div>
                        </details>
                    </div>

                    <button type="submit" class="btn btn-primary">Aplicar</button>
                </div>
            </form>
        <?php endif; ?>
    </div><!-- /toolbar -->

    <?php if ($pagina === 'comparativo'): $dp = $dadosPagina ?? [];
        $modoAtivo = $dp['modoComp'] ?? (in_array($_GET['modo'] ?? '', ['meses','tri','tri_ano','sem'], true) ? $_GET['modo'] : 'meses'); ?>
    <!-- ── FILTROS DO COMPARATIVO: tipo à esquerda, demais à direita com destaque ── -->
    <form method="get" id="formFiltros" class="comp-filtros">
        <input type="hidden" name="pagina" value="<?= htmlspecialchars($pagina) ?>">

        <div class="filtro-grupo">
            <label>Tipo de comparação</label>
            <div class="modo-radio">
                <label><input type="radio" name="modo" value="meses"   <?= $modoAtivo==='meses'  ?'checked':'' ?> onchange="this.form.submit()">Meses vs Anos Anteriores</label>
                <label><input type="radio" name="modo" value="tri"     <?= $modoAtivo==='tri'    ?'checked':'' ?> onchange="this.form.submit()">Trimestres vs Anos Anteriores</label>
                <label><input type="radio" name="modo" value="tri_ano" <?= $modoAtivo==='tri_ano'?'checked':'' ?> onchange="this.form.submit()">Trimestres <?= $hoje->format('Y') ?></label>
                <label><input type="radio" name="modo" value="sem"     <?= $modoAtivo==='sem'    ?'checked':'' ?> onchange="this.form.submit()">Semestre vs Semestre</label>
            </div>
        </div>

        <div class="painel-destaque">
            <div class="filtros">
                <?php if (count($escoposDisponiveis) > 1): ?>
                    <div class="filtro-grupo">
                        <label>Carteira</label>
                        <select name="escopo" onchange="this.form.submit()">
                            <?php foreach ($escoposDisponiveis as $k => $e): ?>
                            <option value="<?= $k ?>" <?= $k === $escopo ? 'selected' : '' ?>><?= $e['label'] ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <?php else: ?>
                    <input type="hidden" name="escopo" value="<?= htmlspecialchars($escopo) ?>">
                    <?php endif; ?>

                <?php if ($modoAtivo === 'meses'): ?>
                <div class="filtro-grupo">
                    <label>Meses</label>
                    <div class="meses-check">
                        <?php foreach (TC_MESES_ABREV as $m => $ab): ?>
                        <label><input type="checkbox" name="meses[]" value="<?= $m ?>" <?= in_array($m, $dp['mesesSel'] ?? [], true)?'checked':'' ?>><?= $ab ?></label>
                        <?php endforeach; ?>
                    </div>
                </div>
                <?php elseif ($modoAtivo === 'tri'): ?>
                <div class="filtro-grupo">
                    <label>Trimestre</label>
                    <select name="tri">
                        <?php foreach ([1=>'T1 (Jan-Mar)',2=>'T2 (Abr-Jun)',3=>'T3 (Jul-Set)',4=>'T4 (Out-Dez)'] as $t => $lb): ?>
                        <option value="<?= $t ?>" <?= ($dp['triSel'] ?? 1)===$t?'selected':'' ?>><?= $lb ?></option>
                        <?php endforeach; ?>
                    </select>
                </div>
                <?php elseif ($modoAtivo === 'sem'):
                    // Anos e defaults do seletor (robustos mesmo sem dados construídos)
                    $anoHoje  = (int)$hoje->format('Y');
                    $semHoje  = ((int)$hoje->format('n') <= 6) ? 1 : 2;
                    $anosSemUI = $dp['anosOpc'] ?? [$anoHoje, $anoHoje - 1, $anoHoje - 2];
                    $sB = $dp['semBaseSel']  ?? (in_array((int)($_GET['sbase'] ?? 0),  [1,2], true)      ? (int)$_GET['sbase']  : ($semHoje===1?2:1));
                    $aB = $dp['anoBaseSel']  ?? (in_array((int)($_GET['abase'] ?? 0),  $anosSemUI, true) ? (int)$_GET['abase']  : ($semHoje===1?$anoHoje-1:$anoHoje));
                    $sA = $dp['semAtualSel'] ?? (in_array((int)($_GET['satual'] ?? 0), [1,2], true)      ? (int)$_GET['satual'] : $semHoje);
                    $aA = $dp['anoAtualSel'] ?? (in_array((int)($_GET['aatual'] ?? 0), $anosSemUI, true) ? (int)$_GET['aatual'] : $anoHoje);
                ?>
                <div class="filtro-grupo">
                    <label>Semestre base</label>
                    <div class="sem-par">
                        <select name="sbase">
                            <option value="1" <?= $sB===1?'selected':'' ?>>1º Sem</option>
                            <option value="2" <?= $sB===2?'selected':'' ?>>2º Sem</option>
                        </select>
                        <select name="abase">
                            <?php foreach ($anosSemUI as $an): ?>
                            <option value="<?= $an ?>" <?= $aB===$an?'selected':'' ?>><?= $an ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                </div>
                <div class="filtro-grupo">
                    <label>Semestre comparado</label>
                    <div class="sem-par">
                        <select name="satual">
                            <option value="1" <?= $sA===1?'selected':'' ?>>1º Sem</option>
                            <option value="2" <?= $sA===2?'selected':'' ?>>2º Sem</option>
                        </select>
                        <select name="aatual">
                            <?php foreach ($anosSemUI as $an): ?>
                            <option value="<?= $an ?>" <?= $aA===$an?'selected':'' ?>><?= $an ?></option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                </div>
                <?php endif; ?>

                <div class="filtro-grupo">
                    <label>Vendedor (RCA)</label>
                    <details class="rca-drop">
                        <summary id="rcaResumo">
                            <?= !$rcasSel ? 'Todos' : (count($rcasSel) === 1 ? htmlspecialchars($rcasSel[0]) : count($rcasSel) . ' selecionados') ?>
                        </summary>
                        <div class="rca-lista">
                            <?php foreach ($rcaOpts as $nome): ?>
                            <label><input type="checkbox" name="rca[]" value="<?= htmlspecialchars($nome) ?>" <?= in_array($nome, $rcasSel, true)?'checked':'' ?>> <?= htmlspecialchars($nome) ?></label>
                            <?php endforeach; ?>
                        </div>
                    </details>
                </div>

                <button type="submit" class="btn btn-primary">Aplicar</button>
            </div>
        </div>
    </form>
    <?php endif; ?>
    <?php if ($erroConexao): ?>
        <div class="warn-box">⚠️ Erro de conexão: <code><?= htmlspecialchars($erroConexao) ?></code></div>
    <?php elseif (!$raw && !$rawComp): ?>
        <div class="info-box">ℹ️ Nenhum dado encontrado para o período selecionado.</div>
    <?php endif; ?>

    <?php foreach ($avisos as [$tipo, $msg]): ?>
        <div class="<?= $tipo === 'warn' ? 'warn-box' : 'info-box' ?>"><?= $msg ?></div>
    <?php endforeach; ?>

<?php // ═══════════════════ PÁGINA: CLIENTES ALVO ═══════════════════
if ($pagina === 'alvo' && $dadosPagina): extract($dadosPagina); ?>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Clientes em queda</div>
            <div class="kpi-value"><?= tc_fmt_num(count($res)) ?></div>
            <div class="kpi-sub">compraram, mas abaixo da referência</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Vendedores Afetados</div>
            <div class="kpi-value"><?= tc_fmt_num($rcasAfetados) ?></div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Variação Média</div>
            <div class="kpi-value"><?= number_format($varMedia, 1, ',', '.') ?>%</div>
            <div class="kpi-sub"><?= $modoTri ? 'no trimestre mais recente' : 'no mês mais recente' ?></div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Perda Estimada</div>
            <div class="kpi-value"><?= tc_fmt_brl($perdaTotal) ?></div>
            <div class="kpi-sub"><?= htmlspecialchars($labelPerda) ?></div>
        </div>
    </div>

    <div class="card">
        <div class="card-header"><h2>📈 <?= htmlspecialchars($tituloGraf) ?></h2></div>
        <div class="card-body"><div id="chartQueda" style="min-height:300px"></div></div>
    </div>

    <div class="info-box">📐 <b>Lógica:</b> <?= $txtLogica ?></div>

    <p class="section-title">Detalhamento por Cliente</p>
    <div class="data-table-wrapper">
        <table class="data-table" id="tblAlvo">
            <thead><tr>
                <th data-t="n">Cód <span class="seta">↕</span></th>
                <th data-t="s">Cliente <span class="seta">↕</span></th>
                <th data-t="s">Cidade <span class="seta">↕</span></th>
                <th data-t="n">Cód RCA <span class="seta">↕</span></th>
                <th data-t="s">RCA <span class="seta">↕</span></th>
                <?php foreach ($mesesHist as $m): ?><th data-t="n"><?= htmlspecialchars($m) ?> <span class="seta">↕</span></th><?php endforeach; ?>
                <?php if ($labelMedia): ?><th data-t="n">Média <span class="seta">↕</span></th><?php endif; ?>
                <th data-t="n"><?= htmlspecialchars($mesAtual) ?> <span class="seta">↕</span></th>
                <th data-t="n">Var % <span class="seta">↕</span></th>
            </tr></thead>
            <tbody>
            <?php foreach ($res as $r): ?>
                <tr>
                    <td class="td-num" data-v="<?= $r['COD_CLI'] ?>"><?= $r['COD_CLI'] ?></td>
                    <td class="td-cliente" title="<?= htmlspecialchars($r['NOME_CLIENTE']) ?>">
                        <?= htmlspecialchars($r['NOME_CLIENTE']) ?>
                        <?php if ($r['NUNCA_COMPROU']): ?> <span class="badge-nunca">nunca comprou</span><?php endif; ?>
                    </td>
                    <td class="td-cidade" title="<?= htmlspecialchars($r['CIDADE']) ?>"><?= htmlspecialchars($r['CIDADE']) ?></td>
                    <td class="td-codrca" data-v="<?= $r['COD_RCA'] ?>"><?= $r['COD_RCA'] ?></td>
                    <td class="td-rca" title="<?= htmlspecialchars($r['NOME_RCA']) ?>"><?= htmlspecialchars($r['NOME_RCA']) ?></td>
                    <?php foreach ($mesesHist as $m): ?>
                    <td class="td-num" data-v="<?= $r['vals'][$m] ?>"><?= tc_fmt_brl2($r['vals'][$m]) ?></td>
                    <?php endforeach; ?>
                    <?php if ($labelMedia): ?>
                    <td class="td-num" data-v="<?= $r['media'] ?>"><?= tc_fmt_brl2($r['media']) ?></td>
                    <?php endif; ?>
                    <td class="td-num" data-v="<?= $r['atual'] ?>"><?= tc_fmt_brl2($r['atual']) ?></td>
                    <td class="td-num <?= $r['var'] < 0 ? 'negative-value' : '' ?>" data-v="<?= $r['var'] ?>"><?= number_format($r['var'], 1, ',', '.') ?>%</td>
                </tr>
            <?php endforeach; ?>
            </tbody>
        </table>
    </div>

    <p class="section-title">Exportar</p>
    <button class="btn" onclick="exportarExcel()">⬇️ Baixar Excel</button>

    <script>
    new ApexCharts(document.querySelector('#chartQueda'), {
        chart: { type: 'line', height: 300, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
        series: [
            { name: '<?= (int)$anoAtualLbl ?>',     data: <?= json_encode($yAtual) ?> },
            { name: '<?= (int)$anoAtualLbl - 1 ?>', data: <?= json_encode($yAnterior) ?> }
        ],
        colors: ['#EF4444', '#2563EB'],
        stroke: { width: [2.5, 2.5], curve: 'smooth', dashArray: [0, 6] },
        dataLabels: {
            enabled: true,
            formatter: v => 'R$ ' + Math.round(v).toLocaleString('pt-BR'),
            offsetY: -7,
            style: { fontSize: '10px', fontWeight: 600 },
            background: { enabled: true, borderRadius: 4, padding: 4, opacity: 0.9 }
        },
        markers: { size: 5, strokeWidth: 2, strokeColors: ['#EF4444', '#2563EB'], colors: ['#fff'] },
        xaxis: { categories: <?= json_encode($labelsX) ?>, labels: { style: { colors: '#374151' } } },
        yaxis: { labels: { formatter: v => 'R$ ' + Math.round(v).toLocaleString('pt-BR') } },
        grid: { borderColor: '#F3F4F6' },
        legend: { position: 'top', horizontalAlign: 'left' },
        tooltip: { y: { formatter: v => 'R$ ' + Math.round(v).toLocaleString('pt-BR') } }
    }).render();
    </script>

<?php // ═══════════════════ PÁGINA: CLIENTES SEM COMPRAS ═══════════════════
elseif ($pagina === 'sem' && $dadosPagina): extract($dadosPagina); ?>

    <div class="info-box">📅 Analisando <b><?= $diasPeriodo ?> dias</b> (<?= $dtIni->format('d/m/Y') ?> a <?= $dtFim->format('d/m/Y') ?>)</div>

    <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr)">
        <div class="kpi-card">
            <div class="kpi-label">Clientes sem compras (&lt;365 dias)</div>
            <div class="kpi-value"><?= tc_fmt_num(count($semCompra)) ?></div>
            <div class="kpi-sub"><?= number_format($pct, 1, ',', '.') ?>% da carteira (<?= tc_fmt_num($novosSemPedido) ?> nunca compraram)</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Vendedores Afetados</div>
            <div class="kpi-value"><?= tc_fmt_num($rcasAfetados) ?></div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Média Dias s/ Compra</div>
            <div class="kpi-value"><?= number_format($mediaDias, 0, ',', '.') ?></div>
            <div class="kpi-sub">dias desde a última compra (exclui quem nunca comprou)</div>
        </div>
    </div>

    <p class="section-title">Clientes sem compras no último ano</p>
    <div class="data-table-wrapper">
        <table class="data-table" id="tblSem">
            <thead><tr>
                <th data-t="n">Código <span class="seta">↕</span></th>
                <th data-t="s">Cliente <span class="seta">↕</span></th>
                <th data-t="s">Cidade <span class="seta">↕</span></th>
                <th data-t="n">Cód RCA <span class="seta">↕</span></th>
                <th data-t="s">RCA <span class="seta">↕</span></th>
                <th data-t="n">Dias Inativo <span class="seta">↕</span></th>
            </tr></thead>
            <tbody>
            <?php foreach ($semCompra as $r): ?>
                <tr>
                    <td class="td-num" data-v="<?= $r['COD_CLI'] ?>"><?= $r['COD_CLI'] ?></td>
                    <td class="td-cliente" title="<?= htmlspecialchars($r['NOME_CLIENTE']) ?>"><?= htmlspecialchars($r['NOME_CLIENTE']) ?></td>
                    <td class="td-cidade" title="<?= htmlspecialchars($r['CIDADE']) ?>"><?= htmlspecialchars($r['CIDADE']) ?></td>
                    <td class="td-codrca" data-v="<?= $r['COD_RCA'] ?>"><?= $r['COD_RCA'] ?></td>
                    <td class="td-rca" title="<?= htmlspecialchars($r['NOME_RCA']) ?>"><?= htmlspecialchars($r['NOME_RCA']) ?></td>
                    <td class="td-num" data-v="<?= $r['nunca'] ? -1 : $r['dias'] ?>">
                        <?= $r['nunca'] ? '<span class="badge-nunca">Nunca comprou</span>' : $r['dias'] ?>
                    </td>
                </tr>
            <?php endforeach; ?>
            </tbody>
        </table>
    </div>

    <p class="section-title">Exportar</p>
    <button class="btn" onclick="exportarExcel()">⬇️ Baixar Excel</button>

<?php // ═══════════════════ PÁGINA: ÍNDICE DE RISCO ═══════════════════
elseif ($pagina === 'risco' && $dadosPagina): extract($dadosPagina); ?>

    <div class="info-box">📅 Período considerado: <b><?= $dtIni->format('d/m/Y') ?></b> a <b><?= $dtFim->format('d/m/Y') ?></b></div>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Frequência Média Esperada</div>
            <div class="kpi-value"><?= number_format($mediaP80, 0, ',', '.') ?> dias</div>
            <div class="kpi-sub">intervalo médio entre compras</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Clientes em Risco</div>
            <div class="kpi-value"><?= tc_fmt_num(count($risco)) ?></div>
            <div class="kpi-sub">Ratio &gt; 1.3</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Ratio Médio</div>
            <div class="kpi-value"><?= number_format($mediaRatio, 1, ',', '.') ?></div>
            <div class="kpi-sub">quanto maior, pior</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Total em Risco</div>
            <div class="kpi-value"><?= tc_fmt_brl($somaTotal) ?></div>
            <div class="kpi-sub">últimos 365 dias</div>
        </div>
    </div>

    <p class="section-title">Clientes com Risco de Abandono</p>
    <div class="data-table-wrapper">
        <table class="data-table" id="tblRisco">
            <thead><tr>
                <th data-t="n">Código <span class="seta">↕</span></th>
                <th data-t="s">Cliente <span class="seta">↕</span></th>
                <th data-t="s">Cidade <span class="seta">↕</span></th>
                <th data-t="n">Cód RCA <span class="seta">↕</span></th>
                <th data-t="s">RCA <span class="seta">↕</span></th>
                <th data-t="n">Freq max Esp <span class="seta">↕</span></th>
                <th data-t="n">Dias Inativo <span class="seta">↕</span></th>
                <th data-t="n">Ind Risco <span class="seta">↕</span></th>
                <th data-t="n">Compra anual <span class="seta">↕</span></th>
            </tr></thead>
            <tbody>
            <?php foreach ($risco as $r): ?>
                <tr>
                    <td class="td-num" data-v="<?= $r['COD_CLI'] ?>"><?= $r['COD_CLI'] ?></td>
                    <td class="td-cliente" title="<?= htmlspecialchars($r['NOME_CLIENTE']) ?>"><?= htmlspecialchars($r['NOME_CLIENTE']) ?></td>
                    <td class="td-cidade" title="<?= htmlspecialchars($r['CIDADE']) ?>"><?= htmlspecialchars($r['CIDADE']) ?></td>
                    <td class="td-codrca" data-v="<?= $r['COD_RCA'] ?>"><?= $r['COD_RCA'] ?></td>
                    <td class="td-rca" title="<?= htmlspecialchars($r['NOME_RCA']) ?>"><?= htmlspecialchars($r['NOME_RCA']) ?></td>
                    <td class="td-num" data-v="<?= $r['p80'] ?>"><?= $r['p80'] ?> dias</td>
                    <td class="td-num" data-v="<?= $r['dias'] ?>"><?= $r['dias'] ?></td>
                    <td class="td-num <?= $r['ratio'] > 1.5 ? 'high-ratio' : ($r['ratio'] > 1 ? 'medium-ratio' : '') ?>" data-v="<?= $r['ratio'] ?>"><?= number_format($r['ratio'], 1, ',', '.') ?></td>
                    <td class="td-num" data-v="<?= $r['total'] ?>"><?= tc_fmt_brl2($r['total']) ?></td>
                </tr>
            <?php endforeach; ?>
            </tbody>
        </table>
    </div>

    <p class="section-title">Exportar</p>
    <button class="btn" onclick="exportarExcel()">⬇️ Baixar Excel</button>

<?php // ═══════════════════ PÁGINA: COMPARATIVO POR PERÍODO ═══════════════════
elseif ($pagina === 'comparativo' && $dadosPagina): extract($dadosPagina); ?>

    <div class="card">
        <div class="card-header"><h2>📈 Faturamento dos clientes em queda por período</h2></div>
        <div class="card-body"><div id="chartComp" style="min-height:280px"></div></div>
    </div>

    <div class="info-box">📐 <b>Lógica:</b> <?= $txtLogicaC ?></div>

    <?php if (count($ordem) < 2): ?>
        <div class="info-box">ℹ️ Ainda não há períodos fechados suficientes para montar a tabela de clientes.</div>
    <?php elseif (!$resComp): ?>
        <div class="info-box">ℹ️ Nenhum cliente encontrado para os filtros e período selecionados.</div>
    <?php else: ?>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Clientes Analisados</div>
            <div class="kpi-value"><?= tc_fmt_num($kpiComp['qtd']) ?></div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Cresceram</div>
            <div class="kpi-value"><?= tc_fmt_num($kpiComp['cresceram']) ?></div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Caíram</div>
            <div class="kpi-value"><?= tc_fmt_num($kpiComp['cairam']) ?></div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Variação Total</div>
            <div class="kpi-value"><?= number_format($kpiComp['variacao'], 1, ',', '.') ?>%</div>
            <div class="kpi-sub"><?= tc_fmt_brl($kpiComp['somaRef']) ?> → <?= tc_fmt_brl($kpiComp['somaAtual']) ?></div>
        </div>
    </div>

    <div class="data-table-wrapper">
        <table class="data-table" id="tblComp">
            <thead><tr>
                <th data-t="n">Cód <span class="seta">↕</span></th>
                <th data-t="s">Cliente <span class="seta">↕</span></th>
                <th data-t="s">Cidade <span class="seta">↕</span></th>
                <th data-t="n">Cód RCA <span class="seta">↕</span></th>
                <th data-t="s">RCA <span class="seta">↕</span></th>
                <?php foreach ($periodosHistC as $p): ?><th data-t="n"><?= htmlspecialchars($p) ?> <span class="seta">↕</span></th><?php endforeach; ?>
                <th data-t="n"><?= htmlspecialchars($periodoAtualC) ?> <span class="seta">↕</span></th>
                <th data-t="n"><?= in_array($modoAtivo, ['tri_ano','sem'], true) ? 'Var %' : 'Var. ano ant.' ?> <span class="seta">↕</span></th>
            </tr></thead>
            <tbody>
            <?php foreach ($resComp as $r): ?>
                <tr>
                    <td class="td-num" data-v="<?= $r['COD_CLI'] ?>"><?= $r['COD_CLI'] ?></td>
                    <td class="td-cliente" title="<?= htmlspecialchars($r['NOME_CLIENTE']) ?>"><?= htmlspecialchars($r['NOME_CLIENTE']) ?></td>
                    <td class="td-cidade" title="<?= htmlspecialchars($r['CIDADE']) ?>"><?= htmlspecialchars($r['CIDADE']) ?></td>
                    <td class="td-codrca" data-v="<?= $r['COD_RCA'] ?>"><?= $r['COD_RCA'] ?></td>
                    <td class="td-rca" title="<?= htmlspecialchars($r['NOME_RCA']) ?>"><?= htmlspecialchars($r['NOME_RCA']) ?></td>
                    <?php foreach ($periodosHistC as $p): ?>
                    <td class="td-num" data-v="<?= $r['vals'][$p] ?>"><?= tc_fmt_brl2($r['vals'][$p]) ?></td>
                    <?php endforeach; ?>
                    <td class="td-num" data-v="<?= $r['atual'] ?>"><?= tc_fmt_brl2($r['atual']) ?></td>
                    <td class="td-num <?= $r['var'] < 0 ? 'negative-value' : '' ?>" data-v="<?= $r['var'] ?>"><?= number_format($r['var'], 1, ',', '.') ?>%</td>
                </tr>
            <?php endforeach; ?>
            </tbody>
        </table>
    </div>

    <p class="section-title">Exportar</p>
    <button class="btn" onclick="exportarExcel()">⬇️ Baixar Excel</button>

    <?php endif; ?>

    <script>
    (function () {
        const labels  = <?= json_encode($ordemGraf) ?>;
        const valores = <?= json_encode(array_map(fn($l) => round($totais[$l], 2), $ordemGraf)) ?>;
        const fechados = <?= count($ordem) ?>;
        const cores = labels.map((_, i) => i < fechados ? '#93C5FD' : '#2563EB');
        new ApexCharts(document.querySelector('#chartComp'), {
            chart: { type: 'bar', height: 280, toolbar: { show: false }, fontFamily: 'Sora, sans-serif' },
            series: [{ name: 'Faturamento', data: valores }],
            plotOptions: { bar: { distributed: true, columnWidth: '48%', borderRadius: 4 } },
            colors: cores,
            dataLabels: {
                enabled: true,
                formatter: v => 'R$ ' + Math.round(v).toLocaleString('pt-BR'),
                offsetY: -20, style: { colors: ['#374151'], fontSize: '11px', fontWeight: 600 }
            },
            xaxis: { categories: labels, labels: { style: { colors: '#374151' } } },
            yaxis: { labels: { formatter: v => 'R$ ' + Math.round(v).toLocaleString('pt-BR') } },
            grid: { borderColor: '#F3F4F6' },
            legend: { show: false },
            tooltip: { y: { formatter: v => 'R$ ' + Math.round(v).toLocaleString('pt-BR') } }
        }).render();
    })();
    </script>

<?php endif; ?>

</div><!-- /page-wrap -->

<?php


// Monta a estrutura que o JS usa pra gerar o Excel (headers, larguras, linhas por vendedor).
$tcExport = null;
if ($dadosPagina) {
    $agrupar = function (array $linhas): array {
        // ordena como o app.py: clientes com histórico primeiro (por código), "nunca comprou" ao final
        usort($linhas, fn($a, $b) => [$a['nunca'] ?? 0, $a['cod']] <=> [$b['nunca'] ?? 0, $b['cod']]);
        $g = [];
        foreach ($linhas as $l) $g[$l['vend']][] = $l['row'];
        ksort($g, SORT_STRING | SORT_FLAG_CASE);
        return $g;
    };

    if ($pagina === 'alvo' || ($pagina === 'comparativo' && !empty($resComp))) {
        $eRes   = $pagina === 'alvo' ? $res : $resComp;
        $eHist  = $pagina === 'alvo' ? $mesesHist : $periodosHistC;
        $eMedia = $pagina === 'alvo' ? $labelMedia : $labelMediaC;
        $eAtual = $pagina === 'alvo' ? $mesAtual : $periodoAtualC;

        $linhas = [];
        foreach ($eRes as $r) {
            $row = [$r['COD_CLI'], $r['NOME_CLIENTE'], $r['CIDADE']];
            foreach ($eHist as $m) $row[] = round($r['vals'][$m], 2);
            if ($eMedia) $row[] = $r['media'];
            $row[] = round($r['atual'], 2);
            $row[] = sprintf('%.1f%%', $r['var']);
            $linhas[] = ['vend' => $r['NOME_RCA'], 'cod' => $r['COD_CLI'], 'nunca' => $r['NUNCA_COMPROU'] ? 1 : 0, 'row' => $row];
        }
        $tcExport = [
            'titulo'  => 'Queda de Vendas',                    // app.py reusa o mesmo export nos dois relatórios
            'sub'     => 'Período: ' . implode('  /  ', array_merge($eHist, [$eAtual])),
            'hdrRow'  => 4,
            'arquivo' => $pagina === 'alvo' ? 'queda_vendas' : 'comparativo_periodo',
            'headers' => array_merge(['CÓD_CLI', 'NOME_CLIENTE', 'CIDADE'], $eHist, $eMedia ? ['Média'] : [], [$eAtual, 'Variação %']),
            'widths'  => ['CÓD CLI' => 11, 'NOME_CLIENTE' => 36, 'Variação %' => 12, 'P80' => 18, 'Média' => 14],
            'vendedores' => $agrupar($linhas),
        ];

    } elseif ($pagina === 'sem') {
        $linhas = [];
        foreach ($semCompra as $r) {
            $linhas[] = ['vend' => $r['NOME_RCA'], 'cod' => $r['COD_CLI'], 'nunca' => $r['nunca'] ? 1 : 0, 'row' => [
                $r['COD_CLI'],
                $r['ultima'] ? (new DateTime($r['ultima']))->format('d/m/Y') : '',
                $r['NOME_CLIENTE'],
                $r['CIDADE'],
                $r['p80'] ?? '',
                $r['nunca'] ? 999 : $r['dias'],
            ]];
        }
        $tcExport = [
            'titulo'  => 'Clientes sem compras',
            'sub'     => '',
            'hdrRow'  => 3,
            'arquivo' => 'sem_compra',
            'headers' => ['CÓD_CLI', 'Última Compra', 'NOME_CLIENTE', 'CIDADE', 'P80', 'Dias Última Compra'],
            'widths'  => ['CÓD CLI' => 11, 'NOME_CLIENTE' => 36, 'Dias Última Compra' => 20, 'P80' => 14, 'Última Compra' => 18],
            'vendedores' => $agrupar($linhas),
        ];

    } elseif ($pagina === 'risco') {
        // mantém a ordem por ratio decrescente (sem reordenar por código), igual ao app.py
        $g = [];
        foreach ($risco as $r) {
            $g[$r['NOME_RCA']][] = [
                $r['COD_CLI'], $r['NOME_CLIENTE'], $r['CIDADE'],
                $r['dias'], $r['p80'], round($r['total'], 2), $r['ratio'],
            ];
        }
        ksort($g, SORT_STRING | SORT_FLAG_CASE);
        $tcExport = [
            'titulo'  => 'Clientes em Risco',
            'sub'     => '',
            'hdrRow'  => 3,
            'arquivo' => 'clientes_risco',
            'headers' => ['CÓD CLI', 'NOME_CLIENTE', 'CIDADE', 'Dias Última Compra', 'P80', 'Total (R$)', 'Ratio'],
            'widths'  => ['CÓD CLI' => 11, 'NOME_CLIENTE' => 36, 'Dias Última Compra' => 20, 'P80' => 13, 'Ratio' => 11, 'Total (R$)' => 17],
            'vendedores' => $g,
        ];
    }
}
?>
<script>
const TC_EXPORT = <?= json_encode($tcExport, JSON_UNESCAPED_UNICODE) ?>;

// ── Exportar Excel — réplica do _write_sheet do app.py (openpyxl) ───────────
function exportarExcel() {
    if (!TC_EXPORT) return;
    const AZUL = '1A1A1A', BRANCO = 'FFFFFF', CINZA = 'F5F5F5';
    const thin  = { style: 'thin', color: { rgb: 'D0D0D0' } };
    const borda = { top: thin, bottom: thin, left: thin, right: thin };
    const nCols   = TC_EXPORT.headers.length;
    const hdrRow  = TC_EXPORT.hdrRow;        // 1-based, como no openpyxl
    const temSub  = !!TC_EXPORT.sub;
    const wb      = XLSX.utils.book_new();
    const usados  = new Set();

    Object.keys(TC_EXPORT.vendedores).forEach(function (vend) {
        const rows = TC_EXPORT.vendedores[vend];
        const ws = {};
        const set = function (r, c, v, s, t) {
            ws[XLSX.utils.encode_cell({ r: r, c: c })] = {
                v: (v === null || v === undefined) ? '' : v,
                t: t || (typeof v === 'number' ? 'n' : 's'),
                s: s
            };
        };

        // Linha 1: título mesclado, fundo escuro
        ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: nCols - 1 } }];
        set(0, 0, TC_EXPORT.titulo + ' — ' + vend, {
            font: { name: 'Inter', bold: true, sz: 12, color: { rgb: BRANCO } },
            fill: { patternType: 'solid', fgColor: { rgb: AZUL } },
            alignment: { horizontal: 'center', vertical: 'center' }
        });
        for (let c = 1; c < nCols; c++) set(0, c, '', { fill: { patternType: 'solid', fgColor: { rgb: AZUL } } });
        const alturas = []; alturas[0] = { hpt: 28 };

        // Linha 2: subtítulo (período) — só na queda/comparativo
        if (temSub) {
            ws['!merges'].push({ s: { r: 1, c: 0 }, e: { r: 1, c: nCols - 1 } });
            set(1, 0, TC_EXPORT.sub, {
                font: { name: 'Inter', italic: true, sz: 9, color: { rgb: '666666' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            });
            alturas[1] = { hpt: 14 }; alturas[2] = { hpt: 4 };
        }

        // Cabeçalho
        const hr = hdrRow - 1;
        TC_EXPORT.headers.forEach(function (h, c) {
            set(hr, c, h, {
                font: { name: 'Inter', bold: true, sz: 10, color: { rgb: BRANCO } },
                fill: { patternType: 'solid', fgColor: { rgb: AZUL } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: borda
            }, 's');
        });
        alturas[hr] = { hpt: 30 };

        // Dados: zebra branco/cinza, tudo centralizado, bordas finas
        rows.forEach(function (rw, i) {
            const rr = hdrRow + i, fill = (i % 2 === 0) ? BRANCO : CINZA;
            rw.forEach(function (v, c) {
                set(rr, c, v, {
                    font: { name: 'Inter', sz: 9 },
                    fill: { patternType: 'solid', fgColor: { rgb: fill } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border: borda
                });
            });
            alturas[rr] = { hpt: 17 };
        });

        const ultLinha = rows.length ? hdrRow + rows.length - 1 : hr;
        ws['!ref']  = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: ultLinha, c: nCols - 1 } });
        ws['!rows'] = alturas;
        ws['!cols'] = TC_EXPORT.headers.map(h => ({ wch: TC_EXPORT.widths[h] || 16 }));
        if (rows.length) {
            ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: hr, c: 0 }, e: { r: ultLinha, c: nCols - 1 } }) };
        }

        // Nome da aba: remove \ / : * ? [ ], corta em 31 e evita duplicidade
        let nome = vend.replace(/[\\\/:\*\?\[\]]/g, '').substring(0, 31).trim() || 'Sheet';
        const base = nome; let sfx = 1;
        while (usados.has(nome)) nome = base.substring(0, 28) + '_' + (sfx++);
        usados.add(nome);
        XLSX.utils.book_append_sheet(wb, ws, nome);
    });

    const d = new Date(), p = n => String(n).padStart(2, '0');
    const ts = d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
    XLSX.writeFile(wb, TC_EXPORT.arquivo + '_' + ts + '.xlsx');
}
</script>
<script>
// ── Ordenação de tabelas (clique no cabeçalho) ──────────────────────────────
document.querySelectorAll('.data-table th').forEach(function (th) {
    th.addEventListener('click', function () {
        const table = th.closest('table');
        const tbody = table.querySelector('tbody');
        const idx   = Array.prototype.indexOf.call(th.parentNode.children, th);
        const tipo  = th.dataset.t || 's';
        const asc   = th.dataset.asc !== '1';
        table.querySelectorAll('th').forEach(o => { if (o !== th) delete o.dataset.asc; });
        th.dataset.asc = asc ? '1' : '0';

        Array.from(tbody.rows)
            .sort(function (a, b) {
                const ca = a.cells[idx], cb = b.cells[idx];
                let va, vb;
                if (tipo === 'n') {
                    va = parseFloat(ca.dataset.v ?? ca.textContent) || 0;
                    vb = parseFloat(cb.dataset.v ?? cb.textContent) || 0;
                } else {
                    va = ca.textContent.trim().toLowerCase();
                    vb = cb.textContent.trim().toLowerCase();
                }
                return (va < vb ? -1 : va > vb ? 1 : 0) * (asc ? 1 : -1);
            })
            .forEach(tr => tbody.appendChild(tr));
    });
});

// ── Resumo do multiselect de RCA ────────────────────────────────────────────
(function () {
    const drop = document.querySelector('.rca-drop');
    if (!drop) return;
    const resumo = document.getElementById('rcaResumo');
    drop.addEventListener('change', function () {
        const sel = Array.from(drop.querySelectorAll('input:checked')).map(i => i.value);
        resumo.textContent = !sel.length ? 'Todos' : (sel.length === 1 ? sel[0] : sel.length + ' selecionados');
    });
    // Fecha ao clicar fora
    document.addEventListener('click', function (e) {
        if (!drop.contains(e.target)) drop.removeAttribute('open');
    });
})();
</script>

</body>
</html>