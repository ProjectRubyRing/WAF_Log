#!/bin/bash
# =============================================================================
#  alb_connlog_to_excel.sh
#
#  ALB (Application Load Balancer) の接続ログ (connection log) の gz ファイルを
#  まとめて解析し、Excel (xlsx) 形式のレポートとして出力します。
#
#   - 入力ディレクトリ配下 (サブディレクトリを含む) の *.gz をすべて読み込みます
#   - 全レコードを日時の昇順に並べ、1 接続 = 1 行で「接続ログ」シートに出力します
#     (ログの各項目を列に分解し、ヘッダに日本語の項目名とフィールド名を表示)
#   - 接続に失敗した行 (tls_verify_status = Failed:～) は赤色で表示します
#   - 件数の集計結果を「サマリー」シートに出力します
#
#  シート構成:
#    サマリー       … 基本情報と各種集計 (接続結果別 / クライアントIP別 / 時間帯別 など)
#    接続ログ       … ログ 1 行 = Excel 1 行の明細 (オートフィルタ・ウィンドウ枠固定付き)
#    ファイル一覧   … 読み込んだ gz ファイルごとの件数・状態
#    解析エラー行   … 形式が不正で解析できなかった行 (該当がある場合のみ)
#    項目説明       … 各列の意味とエラー理由コードの説明
#
#  動作環境: RHEL 9 系 (bash / gawk / gzip / coreutils / findutils / sed)
#            xlsx (ZIP 形式) の作成には zip コマンドを使用し、無い場合は
#            python3 の標準ライブラリ (zipfile) で代替します。
#            追加の Python モジュール等のインストールは不要です。
#
#  使い方: ./alb_connlog_to_excel.sh [オプション] [入力ディレクトリ]
#          詳細は ./alb_connlog_to_excel.sh -h を参照してください。
# =============================================================================

set -euo pipefail
export LC_ALL=C

SCRIPT_NAME="$(basename "$0")"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
readonly SCRIPT_NAME SCRIPT_DIR

# -----------------------------------------------------------------------------
# 設定値 (環境変数で上書き可能)
# -----------------------------------------------------------------------------
# 「接続ログ」シート 1 枚あたりの最大データ行数
# (Excel の上限は 1,048,576 行。超える場合は「接続ログ_2」以降のシートに分割する)
MAX_ROWS_PER_SHEET="${MAX_ROWS_PER_SHEET:-1000000}"
# 「解析エラー行」シートに出力する最大行数
MAX_ERROR_ROWS="${MAX_ERROR_ROWS:-10000}"

# -----------------------------------------------------------------------------
# 共通関数
# -----------------------------------------------------------------------------
usage() {
    cat <<EOF
使い方: ${SCRIPT_NAME} [オプション] [入力ディレクトリ]

ALB の接続ログ (conn_log_*.log.gz) を解析して Excel (xlsx) レポートを出力します。
入力ディレクトリ配下 (サブディレクトリを含む) の *.gz ファイルをすべて読み込みます。

引数:
  入力ディレクトリ       gz ファイルを配置したディレクトリ
                         (省略時: ${SCRIPT_DIR}/input)

オプション:
  -o, --output FILE      出力する Excel ファイルのパス
                         (省略時: ${SCRIPT_DIR}/output/ALB_conn_log_report_YYYYMMDD_HHMMSS.xlsx)
  -z, --tz-offset HOURS  「日時」列に使う UTC からの時差 (時間)。既定値: 9 (JST)
  -n, --top N            サマリーのランキング (クライアントIP 別など) の表示件数。既定値: 50
  -h, --help             このヘルプを表示

環境変数:
  MAX_ROWS_PER_SHEET     「接続ログ」シート 1 枚あたりの最大行数 (既定値: 1000000)
                         超えた分は「接続ログ_2」以降のシートに分割します
  MAX_ERROR_ROWS         「解析エラー行」シートに出力する最大行数 (既定値: 10000)
  TMPDIR                 作業用ディレクトリの作成場所 (既定値: /tmp)
                         ログが大量の場合は空き容量の大きいディレクトリを指定してください

例:
  ${SCRIPT_NAME}
  ${SCRIPT_NAME} ./input
  ${SCRIPT_NAME} -o /tmp/alb_report.xlsx /data/alb_logs
EOF
}

log_info() { printf '[%(%H:%M:%S)T] %s\n' -1 "$*"; }
log_warn() { printf '[%(%H:%M:%S)T] 警告: %s\n' -1 "$*" >&2; }
die()      { printf '[%(%H:%M:%S)T] エラー: %s\n' -1 "$*" >&2; exit 1; }

WORK_DIR=""
cleanup() {
    if [[ -n "$WORK_DIR" && -d "$WORK_DIR" ]]; then
        rm -rf -- "$WORK_DIR"
    fi
}
trap cleanup EXIT
trap 'exit 130' INT TERM

# -----------------------------------------------------------------------------
# 引数の解析
# -----------------------------------------------------------------------------
parse_args() {
    INPUT_DIR=""
    OUTPUT_FILE=""
    TZ_OFFSET="9"
    TOP_N="50"

    while (( $# > 0 )); do
        case "$1" in
            -o|--output)
                (( $# >= 2 )) || die "$1 には出力ファイルのパスを指定してください"
                OUTPUT_FILE="$2"; shift 2 ;;
            -z|--tz-offset)
                (( $# >= 2 )) || die "$1 には UTC からの時差 (時間) を指定してください"
                TZ_OFFSET="$2"; shift 2 ;;
            -n|--top)
                (( $# >= 2 )) || die "$1 には表示件数を指定してください"
                TOP_N="$2"; shift 2 ;;
            -h|--help)
                usage; exit 0 ;;
            --)
                shift; break ;;
            -*)
                die "不明なオプションです: $1 (-h でヘルプを表示します)" ;;
            *)
                [[ -z "$INPUT_DIR" ]] || die "入力ディレクトリは 1 つだけ指定してください"
                INPUT_DIR="$1"; shift ;;
        esac
    done
    if (( $# > 0 )); then
        [[ -z "$INPUT_DIR" && $# -eq 1 ]] || die "入力ディレクトリは 1 つだけ指定してください"
        INPUT_DIR="$1"
    fi

    [[ "$TZ_OFFSET" =~ ^[+-]?[0-9]{1,2}(\.[0-9]+)?$ ]] \
        || die "時差の指定が不正です: ${TZ_OFFSET} (例: 9, 0, -5, 5.5)"
    [[ "$TOP_N" =~ ^[1-9][0-9]*$ ]] \
        || die "表示件数の指定が不正です: ${TOP_N}"
    [[ "$MAX_ROWS_PER_SHEET" =~ ^[1-9][0-9]*$ ]] && (( MAX_ROWS_PER_SHEET <= 1048575 )) \
        || die "MAX_ROWS_PER_SHEET は 1～1048575 の範囲で指定してください: ${MAX_ROWS_PER_SHEET}"
    [[ "$MAX_ERROR_ROWS" =~ ^[0-9]+$ ]] && (( MAX_ERROR_ROWS <= 1048575 )) \
        || die "MAX_ERROR_ROWS は 0～1048575 の範囲で指定してください: ${MAX_ERROR_ROWS}"

    case "$TZ_OFFSET" in
        9|+9|09|+09|9.0|+9.0) TZ_LABEL="JST" ;;
        0|+0|-0|00|0.0)        TZ_LABEL="UTC" ;;
        [+-]*)                 TZ_LABEL="UTC${TZ_OFFSET}" ;;
        *)                     TZ_LABEL="UTC+${TZ_OFFSET}" ;;
    esac

    # 入力ディレクトリ
    INPUT_DIR="${INPUT_DIR:-$SCRIPT_DIR/input}"
    [[ -d "$INPUT_DIR" ]] \
        || die "入力ディレクトリが存在しません: ${INPUT_DIR} (gz ファイルを配置したディレクトリを指定してください)"
    [[ -r "$INPUT_DIR" && -x "$INPUT_DIR" ]] \
        || die "入力ディレクトリを読み取れません: ${INPUT_DIR}"
    INPUT_DIR="$(cd "$INPUT_DIR" && pwd)"

    # 出力ファイル
    if [[ -z "$OUTPUT_FILE" ]]; then
        OUTPUT_FILE="$SCRIPT_DIR/output/ALB_conn_log_report_$(date '+%Y%m%d_%H%M%S').xlsx"
    fi
    [[ "$OUTPUT_FILE" == *.xlsx ]] || OUTPUT_FILE="${OUTPUT_FILE}.xlsx"
    local out_dir
    out_dir="$(dirname "$OUTPUT_FILE")"
    mkdir -p -- "$out_dir" || die "出力ディレクトリを作成できません: ${out_dir}"
    [[ -w "$out_dir" ]] || die "出力ディレクトリに書き込みできません: ${out_dir}"
    OUTPUT_FILE="$(cd "$out_dir" && pwd)/$(basename "$OUTPUT_FILE")"
}

# -----------------------------------------------------------------------------
# 必要なコマンドの確認
# -----------------------------------------------------------------------------
check_requirements() {
    local cmd missing=()
    for cmd in gawk gzip sort find stat mktemp sed head; do
        command -v "$cmd" >/dev/null 2>&1 || missing+=("$cmd")
    done
    (( ${#missing[@]} == 0 )) \
        || die "必要なコマンドが見つかりません: ${missing[*]} (例: dnf install gawk gzip)"

    if command -v zip >/dev/null 2>&1; then
        ZIP_TOOL="zip"
    elif command -v python3 >/dev/null 2>&1; then
        ZIP_TOOL="python3"
    else
        die "xlsx の作成に zip コマンドまたは python3 が必要です (例: dnf install zip)"
    fi
}

# -----------------------------------------------------------------------------
# 作業ディレクトリと、処理に使う定義ファイル・awk プログラムの作成
# -----------------------------------------------------------------------------
prepare_workdir() {
    WORK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/alb_connlog.XXXXXX")" \
        || die "作業ディレクトリを作成できません (TMPDIR を確認してください)"

    AWK_DIR="$WORK_DIR/awk"
    PKG_DIR="$WORK_DIR/pkg"
    mkdir -p "$AWK_DIR" "$PKG_DIR/_rels" "$PKG_DIR/docProps" "$PKG_DIR/xl/_rels" "$PKG_DIR/xl/worksheets"

    PARSED_TSV="$WORK_DIR/parsed.tsv"     # 解析済みレコード
    SORTED_TSV="$WORK_DIR/sorted.tsv"     # 日時順に並べ替えたレコード
    ERRORS_TSV="$WORK_DIR/errors.tsv"     # 解析エラー行
    FILES_TSV="$WORK_DIR/files.tsv"       # ファイルごとの処理結果
    STAT_ONE="$WORK_DIR/stat_one.tsv"     # 1 ファイル分の件数 (作業用)
    GEN_STAT="$WORK_DIR/gen_stat.tsv"     # シート作成結果の件数
    META_FILE="$WORK_DIR/sheets.tsv"      # 作成したシートの一覧 (ファイル名, シート名, フィルタ範囲)
    FIELDS_DEF="$AWK_DIR/fields.def"
    REASONS_DEF="$AWK_DIR/reasons.def"
    : > "$PARSED_TSV"; : > "$ERRORS_TSV"; : > "$FILES_TSV"; : > "$META_FILE"

    write_definitions
    write_awk_programs
    write_styles
}

# 「接続ログ」シートの列定義 (この順に列を出力する)
#   ID|列見出し|ログのフィールド名|列幅|型|説明     ({TZ} はタイムゾーン名に置換)
#   型: int=整数 / decimal=小数 / datetime=日時 / text=文字列
# エラー理由コードの定義
#   コード|日本語の説明|AWS ドキュメント上の説明
write_definitions() {
    cat > "$FIELDS_DEF" <<'DEF'
no|No.|-|8|int|レポート内の通番 (日時の昇順)
dt|日時({TZ})|(timestamp を変換)|23|datetime|timestamp を {TZ} に変換した日時。Excel の日付型のため、フィルタで年・月・日単位の絞り込みができます
ts|タイムスタンプ(UTC)|timestamp|29|text|ロードバランサーが接続の確立に成功または失敗した日時 (ISO 8601 形式・UTC)
cip|クライアントIP|client_ip|17|text|リクエスト元クライアントの IP アドレス
cport|クライアントポート|client_port|12|int|リクエスト元クライアントのポート番号
lport|リスナーポート|listener_port|14|int|リクエストを受信したロードバランサーのリスナーのポート番号
proto|TLSプロトコル|tls_protocol|12|text|[HTTPS リスナー] ハンドシェイクで使用された SSL/TLS プロトコル。SSL/TLS 以外の接続は「-」
cipher|TLS暗号スイート|tls_cipher|30|text|[HTTPS リスナー] ハンドシェイクで使用された暗号スイート。SSL/TLS 以外の接続は「-」
lat|TLSハンドシェイク時間(秒)|tls_handshake_latency|22|decimal|[HTTPS リスナー] ハンドシェイクの確立に要した時間 (秒・ミリ秒精度)。SSL/TLS 以外の接続、またはハンドシェイクに失敗した場合は「-」
result|接続結果|(tls_verify_status から判定)|9|text|tls_verify_status が Success の場合は「成功」、Failed:～ の場合は「失敗」。失敗した接続の行は赤色で表示します
status|TLS検証ステータス|tls_verify_status|26|text|接続要求のステータス。成功時は Success、失敗時は Failed:<エラーコード>
reason|失敗理由(説明)|(エラーコードから変換)|36|text|Failed:<エラーコード> のエラーコードの日本語説明 (下の「エラー理由コード」一覧を参照)
subj|クライアント証明書サブジェクト|leaf_client_cert_subject|40|text|[HTTPS リスナー / mTLS] リーフクライアント証明書のサブジェクト名。SSL/TLS 以外の接続、mTLS が無効、証明書を読み取れない場合は「-」
valid|クライアント証明書有効期間|leaf_client_cert_validity|44|text|[HTTPS リスナー / mTLS] リーフクライアント証明書の有効期間 (NotBefore / NotAfter、ISO 8601 形式)。「-」になる条件はサブジェクトと同じ
serial|クライアント証明書シリアル番号|leaf_client_cert_serial_number|30|text|[HTTPS リスナー / mTLS] リーフクライアント証明書のシリアル番号。「-」になる条件はサブジェクトと同じ
trace|接続トレースID|conn_trace_id|38|text|接続ごとに一意な ID。アクセスログの conn_trace_id と突き合わせることで、接続とリクエストを関連付けられます (古い形式のログでは空欄)
kex|TLS鍵交換方式|tls_keyexchange|17|text|[HTTPS リスナー] TLS / PQ-TLS のハンドシェイクで使用された鍵交換方式。SSL/TLS 以外の接続は「-」(古い形式のログでは空欄)
elb|ロードバランサー|elb|38|text|ロードバランサーのリソース ID (古い形式のログでは空欄)
node|LBノードIP|ip_address|15|text|接続を処理したロードバランサーノードの IP アドレス。内部向けロードバランサーの場合はプライベート IP アドレス (古い形式のログでは空欄)
extra|追加フィールド|(16 番目以降)|24|text|本スクリプトが未対応の 16 番目以降のフィールド (AWS が今後追加する項目など)。スペース区切りのまま出力します
src|ソースファイル|-|60|text|レコードを読み込んだ gz ファイル (入力ディレクトリからの相対パス)
line|行番号|-|8|int|ソースファイル (解凍後) 内の行番号
DEF

    cat > "$REASONS_DEF" <<'DEF'
ClientCertMaxChainDepthExceeded|クライアント証明書チェーンの最大深度を超えている|The maximum client certificate chain depth has been exceeded
ClientCertMaxSizeExceeded|クライアント証明書の最大サイズを超えている|The maximum client certificate size has been exceeded
ClientCertCrlHit|クライアント証明書が CA によって失効されている (CRL に該当)|Client certificate has been revoked by the CA
ClientCertCrlProcessingError|CRL の処理でエラーが発生した|CRL processing error
ClientCertUntrusted|クライアント証明書が信頼されていない|Client certificate is untrusted
ClientCertNotYetValid|クライアント証明書の有効期間が開始していない|Client certificate is not yet valid
ClientCertExpired|クライアント証明書の有効期限が切れている|Client certificate is expired
ClientCertTypeUnsupported|クライアント証明書の種類がサポートされていない|Client certificate type is unsupported
ClientCertInvalid|クライアント証明書が無効|Client certificate is invalid
ClientCertPurposeInvalid|クライアント証明書の用途 (Purpose) が無効|Client certificate purpose is invalid
ClientCertRejected|カスタムサーバー検証によりクライアント証明書が拒否された|Client certificate is rejected by custom server validation
UnmappedConnectionError|分類されていない実行時の接続エラー|Unmapped runtime connection error
ClientCertIncompatible|クライアント証明書が、選択されたリスナーのセキュリティポリシーと互換性がない|Client certificate is incompatible with the chosen listener security policy
DEF
}

write_awk_programs() {
    # -------------------------------------------------------------------------
    # common.awk : xlsx のワークシート XML を出力するための共通関数
    # -------------------------------------------------------------------------
    cat > "$AWK_DIR/common.awk" <<'AWK'
BEGIN {
    CONVFMT = "%.12g"
    # セルのスタイル番号 (styles.xml の cellXfs の並び順と対応)
    ST_PLAIN = 0; ST_HDR = 1; ST_TXT = 2; ST_INT = 3; ST_DT = 4; ST_DEC = 5
    ST_FAIL_OFFSET = 4          # 接続失敗行のスタイル = 通常スタイル + 4 (6～9)
    ST_TITLE = 10; ST_SEC = 11; ST_LABEL = 12; ST_CNT = 13; ST_PCT = 14; ST_WRAP = 15
    ST_CNT_FAIL = 16
    for (i = 1; i <= 52; i++) COL[i] = col_name(i)
}

function col_name(n,   s) {
    s = ""
    while (n > 0) {
        s = substr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", (n - 1) % 26 + 1, 1) s
        n = int((n - 1) / 26)
    }
    return s
}

function xml_esc(s) {
    if (s ~ /[&<>\001-\010\013\014\016-\037]/) {
        gsub(/&/, "\\&amp;", s)
        gsub(/</, "\\&lt;", s)
        gsub(/>/, "\\&gt;", s)
        gsub(/[\001-\010\013\014\016-\037]/, "", s)
    }
    return s
}

# 文字列セル / 数値セル
function c_str(ref, v, s,   sp) {
    if (v == "") return "<c r=\"" ref "\" s=\"" s "\"/>"
    if (length(v) > 32000) {                    # Excel のセル文字数上限対策
        v = substr(v, 1, 32000)
        sub(/[\200-\377]+$/, "", v)
        v = v "...(省略)"
    }
    sp = (v ~ /^[ \t\n]|[ \t\n]$/) ? " xml:space=\"preserve\"" : ""
    return "<c r=\"" ref "\" s=\"" s "\" t=\"inlineStr\"><is><t" sp ">" xml_esc(v) "</t></is></c>"
}
function c_num(ref, v, s) {
    return "<c r=\"" ref "\" s=\"" s "\"><v>" v "</v></c>"
}

# 次に出力する行 (SR + 1) の c 列目のセル
function cs(c, v, s) { return c_str(COL[c] (SR + 1), v, s) }
function cn(c, v, s) { return c_num(COL[c] (SR + 1), v, s) }

# 型に応じたセル (数値として扱えない値は文字列のまま出力)
function cell_typed(c, v, t, off) {
    if (t == "int" && v ~ /^[0-9]+$/ && length(v) <= 15) return cn(c, v, ST_INT + off)
    if (t == "count" && v ~ /^[0-9]+$/ && length(v) <= 15) return cn(c, v, off ? ST_CNT_FAIL : ST_CNT)
    if (t == "decimal" && v ~ /^[0-9]+(\.[0-9]+)?$/) return cn(c, v, ST_DEC + off)
    if (t == "datetime" && v != "") return cn(c, v, ST_DT + off)
    if (t == "wrap") return cs(c, v, ST_WRAP)
    return cs(c, v, ST_TXT + off)
}

# 1 行出力 (SF: 出力先ファイル, SR: 出力済みの行番号)
function emit(cells, ht,   attr) {
    SR++
    attr = (ht > 0) ? " ht=\"" ht "\" customHeight=\"1\"" : ""
    printf "<row r=\"%d\"%s>%s</row>\n", SR, attr, cells > SF
}

# セクション見出し行 (w 列分を塗りつぶす)
function section(title, w,   i, cells) {
    cells = cs(1, title, ST_SEC)
    for (i = 2; i <= w; i++) cells = cells cs(i, "", ST_SEC)
    emit(cells, 20)
}

# ワークシートの開始 / 終了
#   widths: 列幅 (スペース区切り), xs / ys: 固定する列数 / 行数
function sheet_open(f, widths, xs, ys,   w, n, i, tl, ap, pane, cols) {
    pane = ""
    if (xs > 0 || ys > 0) {
        tl = COL[xs + 1] (ys + 1)
        ap = (xs > 0 && ys > 0) ? "bottomRight" : ((ys > 0) ? "bottomLeft" : "topRight")
        pane = "<pane"
        if (xs > 0) pane = pane " xSplit=\"" xs "\""
        if (ys > 0) pane = pane " ySplit=\"" ys "\""
        pane = pane " topLeftCell=\"" tl "\" activePane=\"" ap "\" state=\"frozen\"/>" \
               "<selection pane=\"" ap "\" activeCell=\"" tl "\" sqref=\"" tl "\"/>"
    }
    n = split(widths, w, " ")
    cols = ""
    for (i = 1; i <= n; i++)
        cols = cols "<col min=\"" i "\" max=\"" i "\" width=\"" w[i] "\" customWidth=\"1\"/>"

    printf "%s\n", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" > f
    printf "%s", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" \
                 " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" > f
    printf "<sheetViews><sheetView workbookViewId=\"0\">%s</sheetView></sheetViews>", pane > f
    if (cols != "") printf "<cols>%s</cols>", cols > f
    printf "<sheetData>\n" > f
}
function sheet_close(f, af) {
    printf "</sheetData>" > f
    if (af != "") printf "<autoFilter ref=\"%s\"/>", af > f
    printf "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" > f
    # 「数値が文字列として保存されています」の警告表示を抑止
    printf "<ignoredErrors><ignoredError sqref=\"A1:AZ1048576\" numberStoredAsText=\"1\"/></ignoredErrors>" > f
    printf "</worksheet>\n" > f
    close(f)
}
AWK

    # -------------------------------------------------------------------------
    # parse.awk : 1 ファイル分の接続ログを TSV に変換する
    #   出力列: 1～15 = ログの各フィールド, 16 = 16 番目以降のフィールド,
    #           17 = ソースファイル, 18 = 行番号
    # -------------------------------------------------------------------------
    cat > "$AWK_DIR/parse.awk" <<'AWK'
BEGIN {
    # スペース区切り。ただしダブルクォートで囲まれた値 (証明書サブジェクト) は 1 項目として扱う
    FPAT = "(\"([^\"\\\\]|\\\\.)*\")|([^ ]+)"
    OFS = "\t"
    TS_RE = "^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(\\.[0-9]+)?Z$"
    src = ENVIRON["RPT_SRC"]
    errfile = ENVIRON["RPT_ERR_FILE"]
    statfile = ENVIRON["RPT_STAT_FILE"]
    total = ok = err = 0
    first = last = ""
}
{
    sub(/\r$/, "")
    if ($0 ~ /^[ \t]*$/) next
    total++
    if ($1 !~ TS_RE) { log_error("timestamp の形式が不正 (接続ログ以外のファイルの可能性)"); next }
    if (NF < 11)     { log_error("項目数が不足 (" NF " 項目)"); next }

    for (i = 1; i <= 15; i++) f[i] = (i <= NF) ? norm($i) : ""
    extra = ""
    for (i = 16; i <= NF; i++) extra = extra (i > 16 ? " " : "") norm($i)

    print f[1], f[2], f[3], f[4], f[5], f[6], f[7], f[8], f[9], f[10], \
          f[11], f[12], f[13], f[14], f[15], extra, src, FNR
    ok++
    if (first == "" || f[1] < first) first = f[1]
    if (last == "" || f[1] > last) last = f[1]
}
END {
    printf "%d\t%d\t%d\t%s\t%s\n", total, ok, err, first, last > statfile
}

# 囲みのダブルクォートを外し、タブを空白に置換する
function norm(s) {
    if (length(s) >= 2 && s ~ /^".*"$/) {
        s = substr(s, 2, length(s) - 2)
        gsub(/\\"/, "\"", s)
    }
    gsub(/\t/, " ", s)
    return s
}

function log_error(reason,   line) {
    err++
    line = $0
    gsub(/\t/, " ", line)
    gsub(/[\200-\377]/, "?", line)          # 不正な文字コードで Excel が開けなくなるのを防ぐ
    if (length(line) > 2000) line = substr(line, 1, 2000) "...(省略)"
    print src, FNR, reason, line >> errfile
}
AWK

    # -------------------------------------------------------------------------
    # gen.awk : 日時順に並べたレコードから「接続ログ」「サマリー」シートを作成する
    # -------------------------------------------------------------------------
    cat > "$AWK_DIR/gen.awk" <<'AWK'
BEGIN {
    FS = "\t"
    pkg       = ENVIRON["RPT_PKG_DIR"]
    meta      = ENVIRON["RPT_META_FILE"]
    max_rows  = ENVIRON["RPT_MAX_ROWS"] + 0
    top_n     = ENVIRON["RPT_TOP_N"] + 0
    tz_label  = ENVIRON["RPT_TZ_LABEL"]
    tz_off    = sprintf("%.0f", ENVIRON["RPT_TZ_OFFSET"] * 3600) + 0
    has_extra = ENVIRON["RPT_HAS_EXTRA"] + 0
    split_sheets = (ENVIRON["RPT_TOTAL_ROWS"] + 0 > max_rows)

    load_columns(ENVIRON["RPT_FIELDS_DEF"])
    load_reasons(ENVIRON["RPT_REASONS_DEF"])

    split("日 月 火 水 木 金 土", WDN, " ")
    # TLS ハンドシェイク時間の分布の区分
    LTH[1] = 0.01; LTH[2] = 0.05; LTH[3] = 0.1; LTH[4] = 0.5; LTH[5] = 1; LTH[6] = 3
    LBL[1] = "0.01 秒未満";          LBL[2] = "0.01 秒 ～ 0.05 秒未満"
    LBL[3] = "0.05 秒 ～ 0.1 秒未満"; LBL[4] = "0.1 秒 ～ 0.5 秒未満"
    LBL[5] = "0.5 秒 ～ 1 秒未満";    LBL[6] = "1 秒 ～ 3 秒未満"
    LBL[7] = "3 秒以上"
    LBL[0] = "記録なし「-」(SSL/TLS 以外の接続、ハンドシェイク失敗など)"

    n = nsucc = nfail = 0
    lat_n = lat_sum = 0
    sheet_no = 0
    printf "summary.xml\tサマリー\t\n" >> meta
    open_detail()
}

{
    if (data_rows >= max_rows) { close_detail(); open_detail() }
    data_rows++
    n++

    status = $11
    if (status == "Success")    { result = "成功"; isfail = 0; nsucc++ }
    else if (status ~ /^Failed/) { result = "失敗"; isfail = 1; nfail++ }
    else                         { result = status; isfail = 0 }

    # 日時: Excel のシリアル値 (日付型) と表示用文字列
    ts_parse($1)
    local = TS_EPOCH + tz_off
    base = strftime("%Y/%m/%d %H:%M:%S %w", local, 1)
    disp = substr(base, 1, 19) "." TS_MS
    if (n == 1) first_disp = disp
    last_disp = disp

    V["no"]     = n
    V["dt"]     = sprintf("%.10f", (local + TS_MS / 1000) / 86400 + 25569)
    V["ts"]     = $1
    V["cip"]    = $2
    V["cport"]  = $3
    V["lport"]  = $4
    V["proto"]  = $5
    V["cipher"] = $6
    V["lat"]    = $7
    V["result"] = result
    V["status"] = status
    V["reason"] = isfail ? status_desc(status) : ""
    V["subj"]   = $8
    V["valid"]  = $9
    V["serial"] = $10
    V["trace"]  = $12
    V["kex"]    = $13
    V["elb"]    = $14
    V["node"]   = $15
    V["extra"]  = $16
    V["src"]    = $17
    V["line"]   = $18

    off = isfail ? ST_FAIL_OFFSET : 0
    cells = ""
    for (i = 1; i <= NC; i++) cells = cells cell_typed(i, V[CID[i]], CTYPE[i], off)
    emit(cells)

    # 集計
    agg("result", result)
    agg("status", status)
    agg("proto",  $5)
    agg("cipher", $6)
    agg("kex",    $13)
    agg("lport",  $4)
    agg("elb",    $14)
    agg("node",   $15)
    agg("ip",     $2)
    agg("day",    substr(base, 1, 10) "(" WDN[substr(base, 21, 1) + 1] ")")
    agg("hour",   substr(base, 1, 13) ":00")
    if (isfail) FAILIP[$2]++

    if ($7 ~ /^[0-9]+(\.[0-9]+)?$/) {
        lv = $7 + 0
        lat_n++
        lat_sum += lv
        if (lat_n == 1 || lv < lat_min) lat_min = lv
        if (lat_n == 1 || lv > lat_max) lat_max = lv
        b = 1
        while (b <= 6 && lv >= LTH[b]) b++
        LB[b]++
    } else {
        LB[0]++
    }
}

END {
    close_detail()
    write_summary()
    printf "%d\t%d\t%d\t%d\n", n, nsucc, nfail, sheet_no > ENVIRON["RPT_GEN_STAT"]
}

# ---- 定義ファイルの読み込み ----
function load_columns(file,   line, a) {
    NC = 0
    while ((getline line < file) > 0) {
        if (line == "" || line ~ /^#/) continue
        split(line, a, "|")
        if (a[1] == "extra" && !has_extra) continue     # 追加フィールドは値がある場合のみ列を出力
        NC++
        CID[NC] = a[1]; CHDR[NC] = a[2]; CFLD[NC] = a[3]; CW[NC] = a[4]; CTYPE[NC] = a[5]
        gsub(/\{TZ\}/, tz_label, CHDR[NC])
    }
    close(file)
    if (NC == 0) { print "列定義を読み込めません: " file > "/dev/stderr"; exit 1 }
}
function load_reasons(file,   line, a) {
    while ((getline line < file) > 0) {
        if (line == "" || line ~ /^#/) continue
        split(line, a, "|")
        RDESC[a[1]] = a[2]
    }
    close(file)
}

# ---- 日時 ----
# 1970-01-01 からの日数
function days_from_civil(y, m, d,   era, yoe, doy, doe) {
    y -= (m <= 2)
    era = int(y / 400)
    yoe = y - era * 400
    doy = int((153 * (m + (m > 2 ? -3 : 9)) + 2) / 5) + d - 1
    doe = yoe * 365 + int(yoe / 4) - int(yoe / 100) + doy
    return era * 146097 + doe - 719468
}
# ISO 8601 (例: 2023-10-04T17:05:15.514108Z) を UNIX 時刻 (秒) とミリ秒 (3 桁の文字列) に分解
# (Excel の表示精度はミリ秒まで。マイクロ秒は四捨五入されず切り捨てて扱う)
function ts_parse(ts,   fr) {
    TS_EPOCH = days_from_civil(substr(ts, 1, 4) + 0, substr(ts, 6, 2) + 0, substr(ts, 9, 2) + 0) * 86400 \
               + substr(ts, 12, 2) * 3600 + substr(ts, 15, 2) * 60 + substr(ts, 18, 2)
    fr = ""
    if (substr(ts, 20, 1) == ".") {
        fr = substr(ts, 21)
        sub(/[^0-9].*$/, "", fr)
    }
    TS_MS = substr(fr "000", 1, 3)
}

# ---- 集計 ----
function agg(cat, k) {
    CNT[cat][k]++
    FAIL[cat][k] += isfail
}
function cat_size(cat) {
    return (cat in CNT) ? length(CNT[cat]) : 0
}
function ratio(c, d) {
    return (d > 0) ? sprintf("%.6f", c / d) : 0
}
function disp_key(k) {
    return (k == "") ? "(記録なし)" : k
}
function status_desc(k,   code) {
    if (k == "Success") return "接続成功"
    if (k ~ /^Failed/) {
        code = k
        sub(/^Failed:?/, "", code)
        return (code in RDESC) ? RDESC[code] : "(説明未登録のエラーコード)"
    }
    if (k == "-") return "記録なし (SSL/TLS 以外の接続など)"
    return ""
}

# ---- 「接続ログ」シート ----
function open_detail(   i, widths, cells, hdr) {
    sheet_no++
    SF = pkg "/xl/worksheets/detail_" sheet_no ".xml"
    SR = 0
    data_rows = 0
    widths = ""
    for (i = 1; i <= NC; i++) widths = widths (i > 1 ? " " : "") CW[i]
    sheet_open(SF, widths, 2, 1)            # No. と日時の列、ヘッダ行を固定
    cells = ""
    for (i = 1; i <= NC; i++) {
        hdr = CHDR[i]
        if (CFLD[i] ~ /^[a-z_]+$/) hdr = hdr "\n" CFLD[i]     # 日本語名 + ログのフィールド名
        cells = cells cs(i, hdr, ST_HDR)
    }
    emit(cells, 42)
}
function close_detail(   name, ref) {
    name = split_sheets ? "接続ログ_" sheet_no : "接続ログ"
    ref = "A1:" COL[NC] SR
    sheet_close(SF, ref)
    printf "detail_%d.xml\t%s\t%s\n", sheet_no, name, ref >> meta
}

# ---- 「サマリー」シート ----
function kv(label, v, t,   c) {
    if (t == "s" || v == "") c = cs(2, (v == "") ? "-" : v, ST_TXT)
    else if (t == "n")       c = cn(2, v, ST_CNT)
    else if (t == "p")       c = cn(2, v, ST_PCT)
    else                     c = cn(2, v, ST_DEC)
    emit(cs(1, label, ST_LABEL) c)
}

# 件数の表 (値 / 件数 / 割合 / うち接続失敗)
#   mode: 並び順 (PROCINFO["sorted_in"]), limit: 表示件数 (0 = 全件)
function sum_table(title, keyhdr, cat, mode, limit,   k, i, w, cells, rest_n, rest_c, rest_f) {
    w = (cat == "status") ? 5 : 4
    section(title, w)
    cells = cs(1, keyhdr, ST_HDR) cs(2, "件数", ST_HDR) cs(3, "割合", ST_HDR) cs(4, "うち接続失敗", ST_HDR)
    if (cat == "status") cells = cells cs(5, "説明", ST_HDR)
    emit(cells)
    if (cat_size(cat) == 0) {
        emit(cs(1, "(データなし)", ST_TXT))
        SR++
        return
    }
    i = rest_n = rest_c = rest_f = 0
    PROCINFO["sorted_in"] = mode
    for (k in CNT[cat]) {
        if (limit > 0 && ++i > limit) {
            rest_n++; rest_c += CNT[cat][k]; rest_f += FAIL[cat][k]
            continue
        }
        cells = cs(1, disp_key(k), ST_TXT) cn(2, CNT[cat][k], ST_CNT) \
                cn(3, ratio(CNT[cat][k], n), ST_PCT) cn(4, FAIL[cat][k] + 0, ST_CNT)
        if (cat == "status") cells = cells cs(5, status_desc(k), ST_TXT)
        emit(cells)
    }
    delete PROCINFO["sorted_in"]
    if (rest_n > 0)
        emit(cs(1, "その他 (" rest_n " 種類)", ST_TXT) cn(2, rest_c, ST_CNT) \
             cn(3, ratio(rest_c, n), ST_PCT) cn(4, rest_f, ST_CNT))
    SR++
}

function fail_ip_table(limit,   k, i, rest_n, rest_c) {
    section("■ 接続失敗 クライアントIP別 (上位 " limit " 件)", 4)
    emit(cs(1, "クライアントIP", ST_HDR) cs(2, "失敗件数", ST_HDR) \
         cs(3, "割合 (失敗全体比)", ST_HDR) cs(4, "総接続件数", ST_HDR))
    if (nfail == 0) {
        emit(cs(1, "(接続失敗なし)", ST_TXT))
        SR++
        return
    }
    i = rest_n = rest_c = 0
    PROCINFO["sorted_in"] = "@val_num_desc"
    for (k in FAILIP) {
        if (++i > limit) { rest_n++; rest_c += FAILIP[k]; continue }
        emit(cs(1, disp_key(k), ST_TXT) cn(2, FAILIP[k], ST_CNT) \
             cn(3, ratio(FAILIP[k], nfail), ST_PCT) cn(4, CNT["ip"][k], ST_CNT))
    }
    delete PROCINFO["sorted_in"]
    if (rest_n > 0)
        emit(cs(1, "その他 (" rest_n " 種類)", ST_TXT) cn(2, rest_c, ST_CNT) \
             cn(3, ratio(rest_c, nfail), ST_PCT) cs(4, "", ST_TXT))
    SR++
}

function latency_table(   b) {
    section("■ TLSハンドシェイク時間の分布", 3)
    emit(cs(1, "区分", ST_HDR) cs(2, "件数", ST_HDR) cs(3, "割合", ST_HDR))
    for (b = 1; b <= 7; b++)
        emit(cs(1, LBL[b], ST_TXT) cn(2, LB[b] + 0, ST_CNT) cn(3, ratio(LB[b], n), ST_PCT))
    emit(cs(1, LBL[0], ST_TXT) cn(2, LB[0] + 0, ST_CNT) cn(3, ratio(LB[0], n), ST_PCT))
    SR++
}

function write_summary() {
    SF = pkg "/xl/worksheets/summary.xml"
    SR = 0
    sheet_open(SF, "52 22 12 14 50", 0, 0)
    emit(cs(1, "ALB 接続ログ 分析レポート", ST_TITLE), 26)
    emit(cs(1, "※ 1 接続ごとの詳細は「接続ログ」シート、各列の意味は「項目説明」シートを参照してください。", ST_PLAIN))
    SR++

    section("■ 基本情報", 2)
    kv("レポート作成日時", ENVIRON["RPT_REPORT_TIME"], "s")
    kv("入力ディレクトリ", ENVIRON["RPT_INPUT_DIR"], "s")
    kv("対象ファイル数", ENVIRON["RPT_FILE_COUNT"], "n")
    kv("　うち解凍エラーのファイル数", ENVIRON["RPT_FILE_NG"], "n")
    kv("ログ行数 (空行を除く)", ENVIRON["RPT_LINES_TOTAL"], "n")
    kv("取込件数 (「接続ログ」シートの行数)", n, "n")
    kv("解析エラー行数 (「解析エラー行」シートを参照)", ENVIRON["RPT_LINES_ERR"], "n")
    if (sheet_no > 1) kv("「接続ログ」シート数 (行数が多いため分割)", sheet_no, "n")
    kv("期間 開始 (" tz_label ")", first_disp, "s")
    kv("期間 終了 (" tz_label ")", last_disp, "s")
    kv("接続成功件数 (Success)", nsucc, "n")
    kv("接続失敗件数 (Failed)", nfail, "n")
    kv("接続失敗率", (n > 0) ? ratio(nfail, n) : "", "p")
    kv("ユニーク クライアントIP 数", cat_size("ip"), "n")
    kv("TLSハンドシェイク時間 最小 (秒)", (lat_n > 0) ? lat_min : "", "d")
    kv("TLSハンドシェイク時間 平均 (秒)", (lat_n > 0) ? sprintf("%.6f", lat_sum / lat_n) : "", "d")
    kv("TLSハンドシェイク時間 最大 (秒)", (lat_n > 0) ? lat_max : "", "d")
    SR++

    sum_table("■ 接続結果別", "接続結果", "result", "@val_num_desc", 0)
    sum_table("■ TLS検証ステータス別", "TLS検証ステータス (tls_verify_status)", "status", "@val_num_desc", 0)
    sum_table("■ TLSプロトコル別", "TLSプロトコル (tls_protocol)", "proto", "@val_num_desc", 0)
    sum_table("■ TLS暗号スイート別", "TLS暗号スイート (tls_cipher)", "cipher", "@val_num_desc", 0)
    sum_table("■ TLS鍵交換方式別", "TLS鍵交換方式 (tls_keyexchange)", "kex", "@val_num_desc", 0)
    sum_table("■ リスナーポート別", "リスナーポート (listener_port)", "lport", "@val_num_desc", 0)
    sum_table("■ クライアントIP別 (上位 " top_n " 件)", "クライアントIP (client_ip)", "ip", "@val_num_desc", top_n)
    fail_ip_table(top_n)
    latency_table()
    sum_table("■ ロードバランサー別", "ロードバランサー (elb)", "elb", "@val_num_desc", 0)
    sum_table("■ LBノードIP別", "LBノードIP (ip_address)", "node", "@val_num_desc", 0)
    sum_table("■ 日別 (" tz_label ")", "日付", "day", "@ind_str_asc", 0)
    sum_table("■ 時間帯別 (" tz_label "・1 時間単位)", "日時", "hour", "@ind_str_asc", 0)

    sheet_close(SF, "")
}
AWK

    # -------------------------------------------------------------------------
    # table.awk : 制御行付きの TSV から表形式のシートを作成する
    #   #T<TAB>文字列     … タイトル行          #S<TAB>文字列 … セクション見出し行
    #   #H<TAB>列1<TAB>…  … 表のヘッダ行        #N<TAB>文字列 … 注記行
    #   #B                … 空行                上記以外      … データ行
    # -------------------------------------------------------------------------
    cat > "$AWK_DIR/table.awk" <<'AWK'
BEGIN {
    FS = "\t"
    SF = ENVIRON["RPT_SHEET_FILE"]
    SR = 0
    ntype = split(ENVIRON["RPT_TYPES"], TY, " ")
    fail_col = ENVIRON["RPT_FAIL_COL"] + 0
    fail_re = ENVIRON["RPT_FAIL_RE"]
    hdr_row = hdr_cols = 0
    sheet_open(SF, ENVIRON["RPT_WIDTHS"], 0, ENVIRON["RPT_FREEZE_ROWS"] + 0)
}
$1 == "#T" { emit(cs(1, $2, ST_TITLE), 24); next }
$1 == "#B" { SR++; next }
$1 == "#N" { emit(cs(1, $2, ST_PLAIN)); next }
$1 == "#S" { section($2, ntype); next }
$1 == "#H" {
    cells = ""
    for (i = 2; i <= NF; i++) cells = cells cs(i - 1, $i, ST_HDR)
    if (hdr_row == 0) { hdr_row = SR + 1; hdr_cols = NF - 1 }
    emit(cells, 20)
    next
}
{
    off = (fail_col > 0 && $fail_col ~ fail_re) ? ST_FAIL_OFFSET : 0
    cells = ""
    for (i = 1; i <= NF; i++) cells = cells cell_typed(i, $i, TY[i], off)
    emit(cells)
}
END {
    ref = ""
    if (ENVIRON["RPT_AUTOFILTER"] == "1" && hdr_row > 0)
        ref = "A" hdr_row ":" COL[hdr_cols] ((SR > hdr_row) ? SR : hdr_row)
    sheet_close(SF, ref)
    fname = SF
    sub(/^.*\//, "", fname)
    printf "%s\t%s\t%s\n", fname, ENVIRON["RPT_SHEET_NAME"], ref >> ENVIRON["RPT_META_FILE"]
}
AWK
}

# xlsx のスタイル定義 (cellXfs の並び順は common.awk の ST_* と対応)
write_styles() {
    cat > "$PKG_DIR/xl/styles.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy/mm/dd hh:mm:ss.000"/><numFmt numFmtId="165" formatCode="0.000"/></numFmts>
<fonts count="5">
<font><sz val="10"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="14"/><color rgb="FF1F4E78"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="10"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><sz val="10"/><color rgb="FF9C0006"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
</fonts>
<fills count="6">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="17">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="1" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="165" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="3" fontId="4" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
XML
}

# -----------------------------------------------------------------------------
# gz ファイルの読み込みと解析
# -----------------------------------------------------------------------------
parse_logs() {
    local -a gz_files
    mapfile -d '' -t gz_files < <(find "$INPUT_DIR" -type f -iname '*.gz' -print0 | sort -z)
    FILE_COUNT=${#gz_files[@]}
    (( FILE_COUNT > 0 )) || die "入力ディレクトリに gz ファイルがありません: ${INPUT_DIR}"
    log_info "対象ファイル: ${FILE_COUNT} 件 (${INPUT_DIR})"

    FILE_NG=0; LINES_TOTAL=0; LINES_OK=0; LINES_ERR=0
    local gz rel size status i=0 f_total f_ok f_err f_first f_last gzip_msg
    local -a rcs
    for gz in "${gz_files[@]}"; do
        i=$((i + 1))
        rel="${gz#"$INPUT_DIR"/}"
        rel="${rel//$'\t'/ }"
        size="$(stat -c %s -- "$gz")"
        : > "$STAT_ONE"

        set +e
        gzip -dc -- "$gz" 2>"$WORK_DIR/gzip_err.txt" \
            | RPT_SRC="$rel" RPT_ERR_FILE="$ERRORS_TSV" RPT_STAT_FILE="$STAT_ONE" \
              gawk -f "$AWK_DIR/parse.awk" >> "$PARSED_TSV"
        rcs=("${PIPESTATUS[@]}")
        set -e
        (( rcs[1] == 0 )) || die "ログの解析処理でエラーが発生しました: ${rel}"

        f_total=0; f_ok=0; f_err=0; f_first=""; f_last=""
        IFS=$'\t' read -r f_total f_ok f_err f_first f_last < "$STAT_ONE" || true

        if (( rcs[0] == 0 )); then
            status="正常"
            (( f_total > 0 )) || status="データなし"
        elif (( rcs[0] == 2 )); then
            status="正常 (gzip の警告あり)"
            log_warn "gzip の警告: ${rel}: $(head -c 300 "$WORK_DIR/gzip_err.txt" | tr '\n' ' ')"
        else
            status="解凍エラー (読み取れた行まで取込)"
            FILE_NG=$((FILE_NG + 1))
            gzip_msg="$(head -c 300 "$WORK_DIR/gzip_err.txt" | tr '\n' ' ')"
            log_warn "gz ファイルを解凍できませんでした (読み取れた行まで取り込みます): ${rel}: ${gzip_msg}"
        fi

        printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
            "$rel" "$status" "$size" "$f_total" "$f_ok" "$f_err" "$f_first" "$f_last" >> "$FILES_TSV"
        LINES_TOTAL=$((LINES_TOTAL + f_total))
        LINES_OK=$((LINES_OK + f_ok))
        LINES_ERR=$((LINES_ERR + f_err))
        log_info "  [${i}/${FILE_COUNT}] ${rel} : 取込 ${f_ok} 行 / 解析エラー ${f_err} 行"
    done

    (( LINES_ERR == 0 )) || log_warn "解析できない行が ${LINES_ERR} 行ありました (「解析エラー行」シートを参照してください)"
}

# -----------------------------------------------------------------------------
# 日時順に並べ替え
# -----------------------------------------------------------------------------
sort_records() {
    log_info "日時順に並べ替えています (${LINES_OK} 件)..."
    sort -t $'\t' -k1,1 -s -T "$WORK_DIR" -o "$SORTED_TSV" "$PARSED_TSV" \
        || die "並べ替えに失敗しました (作業ディレクトリの空き容量を確認してください: ${WORK_DIR})"
    rm -f -- "$PARSED_TSV"

    # 16 番目以降のフィールドが存在するか (存在する場合のみ「追加フィールド」列を出力)
    HAS_EXTRA=0
    if gawk -F'\t' '$16 != "" { found = 1; exit } END { exit !found }' "$SORTED_TSV"; then
        HAS_EXTRA=1
    fi
}

# -----------------------------------------------------------------------------
# 各シートの作成
# -----------------------------------------------------------------------------
# 表形式のシートを作成する
#   $1=シートのファイル名 $2=シート名 $3=入力 TSV $4=列幅 $5=列の型
#   $6=固定する行数 $7=オートフィルタ (1/0) $8=失敗判定の列番号 $9=失敗判定の正規表現
write_table_sheet() {
    RPT_SHEET_FILE="$PKG_DIR/xl/worksheets/$1" RPT_SHEET_NAME="$2" \
    RPT_WIDTHS="$4" RPT_TYPES="$5" RPT_FREEZE_ROWS="$6" RPT_AUTOFILTER="$7" \
    RPT_FAIL_COL="${8:-0}" RPT_FAIL_RE="${9:-}" RPT_META_FILE="$META_FILE" \
        gawk -f "$AWK_DIR/common.awk" -f "$AWK_DIR/table.awk" "$3"
}

generate_sheets() {
    log_info "Excel のシートを作成しています..."

    # サマリー / 接続ログ
    RPT_PKG_DIR="$PKG_DIR" RPT_META_FILE="$META_FILE" RPT_GEN_STAT="$GEN_STAT" \
    RPT_MAX_ROWS="$MAX_ROWS_PER_SHEET" RPT_TOP_N="$TOP_N" \
    RPT_TZ_OFFSET="$TZ_OFFSET" RPT_TZ_LABEL="$TZ_LABEL" \
    RPT_HAS_EXTRA="$HAS_EXTRA" RPT_TOTAL_ROWS="$LINES_OK" \
    RPT_FIELDS_DEF="$FIELDS_DEF" RPT_REASONS_DEF="$REASONS_DEF" \
    RPT_REPORT_TIME="$(date '+%Y/%m/%d %H:%M:%S')" RPT_INPUT_DIR="$INPUT_DIR" \
    RPT_FILE_COUNT="$FILE_COUNT" RPT_FILE_NG="$FILE_NG" \
    RPT_LINES_TOTAL="$LINES_TOTAL" RPT_LINES_ERR="$LINES_ERR" \
        gawk -f "$AWK_DIR/common.awk" -f "$AWK_DIR/gen.awk" "$SORTED_TSV" \
        || die "「接続ログ」シートの作成に失敗しました"

    # ファイル一覧
    {
        printf '#H\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
            "No." "ファイル名 (入力ディレクトリからの相対パス)" "状態" "サイズ(バイト)" \
            "ログ行数" "取込件数" "解析エラー行数" "最初のタイムスタンプ(UTC)" "最後のタイムスタンプ(UTC)"
        gawk -F'\t' -v OFS='\t' '{ print NR, $0 }' "$FILES_TSV"
    } > "$WORK_DIR/sheet_files.tsv"
    write_table_sheet "files.xml" "ファイル一覧" "$WORK_DIR/sheet_files.tsv" \
        "6 100 30 14 10 10 13 29 29" "int text text count count count count text text" \
        1 1 3 "エラー"

    # 解析エラー行 (該当がある場合のみ)
    if [[ -s "$ERRORS_TSV" ]] && (( MAX_ERROR_ROWS > 0 )); then
        {
            printf '#H\t%s\t%s\t%s\t%s\t%s\n' "No." "ソースファイル" "行番号" "エラー内容" "行の内容"
            head -n "$MAX_ERROR_ROWS" "$ERRORS_TSV" | gawk -F'\t' -v OFS='\t' '{ print NR, $0 }'
        } > "$WORK_DIR/sheet_errors.tsv"
        write_table_sheet "errors.xml" "解析エラー行" "$WORK_DIR/sheet_errors.tsv" \
            "6 60 8 40 120" "int text int text text" 1 1
    fi

    # 項目説明
    {
        printf '#T\t%s\n' "項目説明"
        printf '#N\t%s\n' "「接続ログ」シートの各列の意味と、ALB 接続ログのフィールドとの対応です。"
        printf '#B\n'
        printf '#S\t%s\n' "■ 「接続ログ」シートの列"
        printf '#H\t%s\t%s\t%s\n' "列見出し" "ログのフィールド名" "説明"
        gawk -F'|' -v OFS='\t' -v tz="$TZ_LABEL" -v has_extra="$HAS_EXTRA" '
            /^#/ || NF < 6 { next }
            $1 == "extra" && has_extra != 1 { next }
            { h = $2; d = $6; gsub(/\{TZ\}/, tz, h); gsub(/\{TZ\}/, tz, d); print h, $3, d }
        ' "$FIELDS_DEF"
        printf '#B\n'
        printf '#S\t%s\n' "■ エラー理由コード (tls_verify_status が Failed:<エラーコード> の場合)"
        printf '#H\t%s\t%s\t%s\n' "エラーコード" "説明" "AWS ドキュメント上の説明"
        gawk -F'|' -v OFS='\t' '!/^#/ && NF >= 3 { print $1, $2, $3 }' "$REASONS_DEF"
        printf '#B\n'
        printf '#S\t%s\n' "■ 補足"
        printf '#N\t%s\n' "・1 行 = 1 接続です。持続的接続 (Keep-Alive) では、複数のリクエストがあっても接続ログは 1 行になります。"
        printf '#N\t%s\n' "・「接続ログ」シートは日時の昇順です。1 シートあたり ${MAX_ROWS_PER_SHEET} 行を超える場合は「接続ログ_2」以降のシートに分割します。"
        printf '#N\t%s\n' "・接続に失敗した行 (tls_verify_status が Failed:～) は赤色で表示しています。"
        printf '#N\t%s\n' "・値が「-」の項目は、ALB がその項目を記録していないことを表します (SSL/TLS 以外の接続、mTLS が無効な場合など)。"
        printf '#N\t%s\n' "・空欄の項目は、古い形式のログのため、そのフィールド自体が出力されていないことを表します。"
        printf '#N\t%s\n' "・各シートのヘッダ行にはオートフィルタを設定しています。「日時」列は Excel の日付型のため、年・月・日単位で絞り込めます。"
    } > "$WORK_DIR/sheet_fields.tsv"
    write_table_sheet "fields.xml" "項目説明" "$WORK_DIR/sheet_fields.tsv" \
        "30 32 110" "wrap wrap wrap" 0 0
}

# -----------------------------------------------------------------------------
# xlsx パッケージの作成
# -----------------------------------------------------------------------------
build_package() {
    local idx=0 file name ref absref sheets="" rels="" overrides="" defnames=""
    while IFS=$'\t' read -r file name ref; do
        [[ -n "$file" ]] || continue
        idx=$((idx + 1))
        sheets+="<sheet name=\"${name}\" sheetId=\"${idx}\" r:id=\"rId${idx}\"/>"
        rels+="<Relationship Id=\"rId${idx}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/${file}\"/>"
        overrides+="<Override PartName=\"/xl/worksheets/${file}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        if [[ -n "$ref" ]]; then
            absref="$(sed -E 's/([A-Z]+)([0-9]+)/$\1$\2/g' <<< "$ref")"
            defnames+="<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"$((idx - 1))\" hidden=\"1\">'${name}'!${absref}</definedName>"
        fi
    done < "$META_FILE"
    (( idx > 0 )) || die "シートが作成されていません"
    [[ -z "$defnames" ]] || defnames="<definedNames>${defnames}</definedNames>"

    cat > "$PKG_DIR/[Content_Types].xml" <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>${overrides}</Types>
EOF

    cat > "$PKG_DIR/_rels/.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>
EOF

    cat > "$PKG_DIR/docProps/core.xml" <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>ALB 接続ログ 分析レポート</dc:title><dc:creator>${SCRIPT_NAME}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">$(date -u '+%Y-%m-%dT%H:%M:%SZ')</dcterms:created></cp:coreProperties>
EOF

    cat > "$PKG_DIR/xl/workbook.xml" <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets}</sheets>${defnames}</workbook>
EOF

    cat > "$PKG_DIR/xl/_rels/workbook.xml.rels" <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}<Relationship Id="rId$((idx + 1))" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
EOF
}

# ZIP 圧縮して xlsx ファイルを作成する
make_xlsx() {
    local tmp_xlsx="$WORK_DIR/report.xlsx"
    log_info "xlsx ファイルを作成しています (${ZIP_TOOL} を使用)..."
    if [[ "$ZIP_TOOL" == "zip" ]]; then
        (cd "$PKG_DIR" && zip -q -X -r -D "$tmp_xlsx" .) \
            || die "zip コマンドでの xlsx の作成に失敗しました"
    else
        python3 -I - "$PKG_DIR" "$tmp_xlsx" <<'PY' || die "python3 での xlsx の作成に失敗しました"
import os, sys, zipfile
src, out = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.write(os.path.join(src, "[Content_Types].xml"), "[Content_Types].xml")
    for root, dirs, files in os.walk(src):
        dirs.sort()
        for name in sorted(files):
            path = os.path.join(root, name)
            arc = os.path.relpath(path, src).replace(os.sep, "/")
            if arc != "[Content_Types].xml":
                zf.write(path, arc)
PY
    fi
    mv -f -- "$tmp_xlsx" "$OUTPUT_FILE" || die "出力ファイルを書き込めません: ${OUTPUT_FILE}"
}

# -----------------------------------------------------------------------------
# メイン処理
# -----------------------------------------------------------------------------
main() {
    parse_args "$@"
    check_requirements
    prepare_workdir

    log_info "ALB 接続ログの解析を開始します"
    parse_logs
    sort_records
    generate_sheets
    build_package
    make_xlsx

    local rows nsucc nfail sheets
    IFS=$'\t' read -r rows nsucc nfail sheets < "$GEN_STAT"
    log_info "完了しました"
    cat <<EOF

  出力ファイル   : ${OUTPUT_FILE}
  対象ファイル   : ${FILE_COUNT} 件 (うち解凍エラー ${FILE_NG} 件)
  取込件数       : ${rows} 件 (接続成功 ${nsucc} 件 / 接続失敗 ${nfail} 件)
  解析エラー行   : ${LINES_ERR} 件
  接続ログシート : ${sheets} シート
EOF
}

main "$@"
