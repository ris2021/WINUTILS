#define UNICODE
#define _UNICODE

#include <windows.h>
#include <commdlg.h>
#include <richedit.h>
#include <wchar.h>
#include <stdlib.h>
#include <string.h>
#include <stdio.h>

#define ID_OPEN       1001
#define ID_SAVE_RTF   1002
#define OUTER_MARGIN  10
#define INNER_MARGIN  6
#define MAX_COLS      32

#ifndef PFM_STYLE
#define PFM_STYLE 0x00000400
#endif

#define MD_STYLE_NORMAL   (-1)
#define MD_STYLE_HEADING1 (-2)
#define MD_STYLE_HEADING2 (-3)
#define MD_STYLE_HEADING3 (-4)
#define MD_STYLE_HEADING4 (-5)

#ifndef EM_INSERTTABLE
#define EM_INSERTTABLE (WM_USER + 232)
#endif

static HWND g_edit;
static HMODULE g_richedit;

typedef struct {
    const WCHAR *start;
    const WCHAR *end;
} CELL;

typedef struct {
    BOOL bold;
    BOOL italic;
    BOOL strike;
    BOOL code;
} STYLE;

/*
   Local copies of the documented RichEdit table structures.
   This avoids depending on how complete a particular MinGW richedit.h is.
*/
typedef struct {
    BYTE cbRow;
    BYTE cbCell;
    BYTE cCell;
    BYTE cRow;
    LONG dxCellMargin;
    LONG dxIndent;
    LONG dyHeight;

    DWORD nAlignment   : 3;
    DWORD fRTL         : 1;
    DWORD fKeep        : 1;
    DWORD fKeepFollow  : 1;
    DWORD fWrap        : 1;
    DWORD fIdentCells  : 1;
    DWORD reserved     : 24;

    LONG cpStartRow;
    BYTE bTableLevel;
    BYTE iCell;
} MD_TABLEROWPARMS;

typedef struct {
    LONG dxWidth;

    WORD nVertAlign  : 2;
    WORD fMergeTop   : 1;
    WORD fMergePrev  : 1;
    WORD fVertical   : 1;
    WORD fMergeStart : 1;
    WORD fMergeCont  : 1;
    WORD reserved    : 9;

    WORD wShading;

    SHORT dxBrdrLeft;
    SHORT dyBrdrTop;
    SHORT dxBrdrRight;
    SHORT dyBrdrBottom;

    COLORREF crBrdrLeft;
    COLORREF crBrdrTop;
    COLORREF crBrdrRight;
    COLORREF crBrdrBottom;

    COLORREF crBackPat;
    COLORREF crForePat;
} MD_TABLECELLPARMS;


/* -------------------------------------------------------------
   Formatting
   ------------------------------------------------------------- */

static void set_format_ex(HWND edit, int pt, STYLE style)
{
    CHARFORMAT2W cf;

    ZeroMemory(&cf, sizeof(cf));
    cf.cbSize = sizeof(cf);

    cf.dwMask =
        CFM_FACE |
        CFM_SIZE |
        CFM_BOLD |
        CFM_ITALIC |
        CFM_STRIKEOUT |
        CFM_CHARSET |
        CFM_BACKCOLOR;

    cf.yHeight = pt * 20;
    cf.dwEffects = 0;

    if (style.bold)
        cf.dwEffects |= CFE_BOLD;

    if (style.italic)
        cf.dwEffects |= CFE_ITALIC;

    if (style.strike)
        cf.dwEffects |= CFE_STRIKEOUT;

    cf.bCharSet = DEFAULT_CHARSET;

    if (style.code) {
        wcscpy(cf.szFaceName, L"Consolas");
        cf.crBackColor = RGB(240, 240, 240);
    } else {
        wcscpy(cf.szFaceName, L"Arial");
        cf.crBackColor = RGB(255, 255, 255);
    }

    SendMessageW(
        edit,
        EM_SETCHARFORMAT,
        SCF_SELECTION,
        (LPARAM)&cf
    );
}

static void reset_paragraph(HWND edit)
{
    PARAFORMAT2 pf;

    ZeroMemory(&pf, sizeof(pf));
    pf.cbSize = sizeof(pf);

    pf.dwMask =
        PFM_STARTINDENT |
        PFM_RIGHTINDENT |
        PFM_OFFSET |
        PFM_NUMBERING |
        PFM_TABSTOPS |
        PFM_STYLE;

    pf.dxStartIndent = 0;
    pf.dxRightIndent = 0;
    pf.dxOffset = 0;
    pf.wNumbering = 0;
    pf.cTabCount = 0;
    pf.sStyle = MD_STYLE_NORMAL;

    SendMessageW(
        edit,
        EM_SETPARAFORMAT,
        0,
        (LPARAM)&pf
    );
}

static void set_paragraph_style(HWND edit, SHORT style)
{
    PARAFORMAT2 pf;

    ZeroMemory(&pf, sizeof(pf));
    pf.cbSize = sizeof(pf);
    pf.dwMask = PFM_STYLE;
    pf.sStyle = style;

    SendMessageW(
        edit,
        EM_SETPARAFORMAT,
        0,
        (LPARAM)&pf
    );
}

static void set_paragraph_indent(HWND edit, int left_twips, int right_twips)
{
    PARAFORMAT2 pf;

    ZeroMemory(&pf, sizeof(pf));
    pf.cbSize = sizeof(pf);
    pf.dwMask = PFM_STARTINDENT | PFM_RIGHTINDENT;
    pf.dxStartIndent = left_twips;
    pf.dxRightIndent = right_twips;

    SendMessageW(
        edit,
        EM_SETPARAFORMAT,
        0,
        (LPARAM)&pf
    );
}


/* -------------------------------------------------------------
   Text insertion at current selection
   ------------------------------------------------------------- */

static void insert_raw_style(
    HWND edit,
    const WCHAR *text,
    int len,
    int pt,
    STYLE style)
{
    WCHAR *tmp;

    if (len <= 0)
        return;

    set_format_ex(edit, pt, style);

    tmp = (WCHAR *)malloc((len + 1) * sizeof(WCHAR));
    if (!tmp)
        return;

    memcpy(tmp, text, len * sizeof(WCHAR));
    tmp[len] = 0;

    SendMessageW(
        edit,
        EM_REPLACESEL,
        FALSE,
        (LPARAM)tmp
    );

    free(tmp);
}

static void insert_raw(
    HWND edit,
    const WCHAR *text,
    int len,
    int pt,
    BOOL bold)
{
    STYLE s = {0};
    s.bold = bold;
    insert_raw_style(edit, text, len, pt, s);
}

static void insert_inline_markdown(
    HWND edit,
    const WCHAR *start,
    const WCHAR *end,
    int pt,
    BOOL base_bold)
{
    const WCHAR *p = start;
    const WCHAR *chunk = p;

    STYLE style = {0};
    style.bold = base_bold;

    while (p < end) {
        int marker_len = 0;

        enum {
            MARK_NONE,
            MARK_BOLD,
            MARK_ITALIC,
            MARK_BOLDITALIC,
            MARK_STRIKE,
            MARK_CODE
        } mark = MARK_NONE;

        if (!style.code) {
            if (p + 2 < end &&
                ((p[0] == L'*' && p[1] == L'*' && p[2] == L'*') ||
                 (p[0] == L'_' && p[1] == L'_' && p[2] == L'_'))) {
                mark = MARK_BOLDITALIC;
                marker_len = 3;
            }
            else if (p + 1 < end &&
                     ((p[0] == L'*' && p[1] == L'*') ||
                      (p[0] == L'_' && p[1] == L'_'))) {
                mark = MARK_BOLD;
                marker_len = 2;
            }
            else if (p + 1 < end &&
                     p[0] == L'~' && p[1] == L'~') {
                mark = MARK_STRIKE;
                marker_len = 2;
            }
            else if (*p == L'*' || *p == L'_') {
                mark = MARK_ITALIC;
                marker_len = 1;
            }
            else if (*p == L'`') {
                mark = MARK_CODE;
                marker_len = 1;
            }
        }
        else if (*p == L'`') {
            mark = MARK_CODE;
            marker_len = 1;
        }

        if (mark != MARK_NONE) {
            if (p > chunk) {
                insert_raw_style(
                    edit,
                    chunk,
                    (int)(p - chunk),
                    pt,
                    style
                );
            }

            switch (mark) {
            case MARK_BOLD:
                style.bold = !style.bold;
                break;

            case MARK_ITALIC:
                style.italic = !style.italic;
                break;

            case MARK_BOLDITALIC:
                style.bold = !style.bold;
                style.italic = !style.italic;
                break;

            case MARK_STRIKE:
                style.strike = !style.strike;
                break;

            case MARK_CODE:
                style.code = !style.code;
                break;

            default:
                break;
            }

            p += marker_len;
            chunk = p;
            continue;
        }

        p++;
    }

    if (p > chunk) {
        insert_raw_style(
            edit,
            chunk,
            (int)(p - chunk),
            pt,
            style
        );
    }
}


/* -------------------------------------------------------------
   Append normal text at end
   ------------------------------------------------------------- */

static void select_end(HWND edit)
{
    LONG end = GetWindowTextLengthW(edit);
    SendMessageW(edit, EM_SETSEL, end, end);
}

static void append_normal_line(
    HWND edit,
    const WCHAR *start,
    const WCHAR *end,
    int pt,
    BOOL bold)
{
    select_end(edit);
    reset_paragraph(edit);

    insert_inline_markdown(
        edit,
        start,
        end,
        pt,
        bold
    );

    insert_raw(
        edit,
        L"\r\n",
        2,
        pt,
        bold
    );
}

static void append_heading_line(
    HWND edit,
    const WCHAR *start,
    const WCHAR *end,
    int level)
{
    static const SHORT styles[4] = {
        MD_STYLE_HEADING1,
        MD_STYLE_HEADING2,
        MD_STYLE_HEADING3,
        MD_STYLE_HEADING4
    };

    int pt = 15 - level;

    select_end(edit);
    reset_paragraph(edit);
    set_paragraph_style(edit, styles[level - 1]);

    insert_inline_markdown(
        edit,
        start,
        end,
        pt,
        TRUE
    );

    insert_raw(
        edit,
        L"\r\n",
        2,
        pt,
        TRUE
    );
}


/* -------------------------------------------------------------
   Line helpers
   ------------------------------------------------------------- */

static const WCHAR *line_end(const WCHAR *p)
{
    while (*p && *p != L'\r' && *p != L'\n')
        p++;

    return p;
}

static const WCHAR *next_line(const WCHAR *p)
{
    p = line_end(p);

    if (*p == L'\r')
        p++;

    if (*p == L'\n')
        p++;

    return p;
}

static BOOL is_blank_line(
    const WCHAR *start,
    const WCHAR *end)
{
    while (start < end) {
        if (*start != L' ' &&
            *start != L'\t') {
            return FALSE;
        }

        start++;
    }

    return TRUE;
}

static BOOL is_horizontal_rule(
    const WCHAR *start,
    const WCHAR *end)
{
    WCHAR ch = 0;
    int count = 0;

    while (start < end &&
           (*start == L' ' ||
            *start == L'\t')) {
        start++;
    }

    while (end > start &&
           (end[-1] == L' ' ||
            end[-1] == L'\t')) {
        end--;
    }

    while (start < end) {
        if (*start == L' ' ||
            *start == L'\t') {
            start++;
            continue;
        }

        if (*start != L'-' &&
            *start != L'*' &&
            *start != L'_') {
            return FALSE;
        }

        if (!ch)
            ch = *start;
        else if (*start != ch)
            return FALSE;

        count++;
        start++;
    }

    return count >= 3;
}


/* -------------------------------------------------------------
   Quotes / lists / rule
   ------------------------------------------------------------- */

static void append_quote_line(
    HWND edit,
    const WCHAR *start,
    const WCHAR *end)
{
    STYLE s = {0};
    s.italic = TRUE;

    while (start < end &&
           (*start == L' ' ||
            *start == L'\t')) {
        start++;
    }

    if (start < end && *start == L'>')
        start++;

    if (start < end && *start == L' ')
        start++;

    select_end(edit);
    reset_paragraph(edit);
    set_paragraph_indent(edit, 360, 120);

    insert_raw_style(
        edit,
        L"\x201C",
        1,
        10,
        s
    );

    insert_inline_markdown(
        edit,
        start,
        end,
        10,
        FALSE
    );

    insert_raw_style(
        edit,
        L"\x201D\r\n",
        3,
        10,
        s
    );
}

static BOOL parse_unordered_list(
    const WCHAR *start,
    const WCHAR *end,
    const WCHAR **content)
{
    const WCHAR *p = start;

    while (p < end &&
           (*p == L' ' ||
            *p == L'\t')) {
        p++;
    }

    if (p + 1 < end &&
        (*p == L'-' ||
         *p == L'*' ||
         *p == L'+') &&
        (p[1] == L' ' ||
         p[1] == L'\t')) {

        *content = p + 2;
        return TRUE;
    }

    return FALSE;
}

static BOOL parse_ordered_list(
    const WCHAR *start,
    const WCHAR *end,
    const WCHAR **content,
    int *number)
{
    const WCHAR *p = start;
    int n = 0;
    int digits = 0;

    while (p < end &&
           (*p == L' ' ||
            *p == L'\t')) {
        p++;
    }

    while (p < end &&
           *p >= L'0' &&
           *p <= L'9') {

        n = n * 10 + (*p - L'0');
        p++;
        digits++;
    }

    if (digits > 0 &&
        p + 1 < end &&
        (*p == L'.' ||
         *p == L')') &&
        (p[1] == L' ' ||
         p[1] == L'\t')) {

        *number = n;
        *content = p + 2;

        return TRUE;
    }

    return FALSE;
}

static void append_list_line(
    HWND edit,
    const WCHAR *content,
    const WCHAR *end,
    BOOL ordered,
    int number)
{
    WCHAR prefix[32];

    select_end(edit);
    reset_paragraph(edit);
    set_paragraph_indent(edit, 300, 0);

    if (ordered) {
        swprintf(
            prefix,
            sizeof(prefix) / sizeof(prefix[0]),
            L"%d. ",
            number
        );

        insert_raw(
            edit,
            prefix,
            (int)wcslen(prefix),
            10,
            FALSE
        );
    }
    else {
        insert_raw(
            edit,
            L"\x2022 ",
            2,
            10,
            FALSE
        );
    }

    insert_inline_markdown(
        edit,
        content,
        end,
        10,
        FALSE
    );

    insert_raw(
        edit,
        L"\r\n",
        2,
        10,
        FALSE
    );
}

static void append_horizontal_rule(HWND edit)
{
    static const WCHAR rule[] =
        L"\x2500\x2500\x2500\x2500\x2500"
        L"\x2500\x2500\x2500\x2500\x2500"
        L"\x2500\x2500\x2500\x2500\x2500"
        L"\x2500\x2500\x2500\x2500\x2500"
        L"\x2500\x2500\x2500\x2500\x2500"
        L"\x2500\x2500\x2500\x2500\x2500"
        L"\r\n";

    select_end(edit);
    reset_paragraph(edit);

    insert_raw(
        edit,
        rule,
        (int)wcslen(rule),
        10,
        FALSE
    );
}


/* -------------------------------------------------------------
   Markdown table parsing
   ------------------------------------------------------------- */

static void trim_cell(CELL *cell)
{
    while (cell->start < cell->end &&
           (*cell->start == L' ' ||
            *cell->start == L'\t')) {
        cell->start++;
    }

    while (cell->end > cell->start &&
           (cell->end[-1] == L' ' ||
            cell->end[-1] == L'\t')) {
        cell->end--;
    }
}

static BOOL line_has_pipe(
    const WCHAR *start,
    const WCHAR *end)
{
    while (start < end) {
        if (*start == L'|')
            return TRUE;

        start++;
    }

    return FALSE;
}

static int split_cells(
    const WCHAR *start,
    const WCHAR *end,
    CELL *cells,
    int max_cells)
{
    int count = 0;
    const WCHAR *cell_start;
    const WCHAR *p;

    while (start < end &&
           (*start == L' ' ||
            *start == L'\t')) {
        start++;
    }

    while (end > start &&
           (end[-1] == L' ' ||
            end[-1] == L'\t')) {
        end--;
    }

    if (start < end && *start == L'|')
        start++;

    if (end > start && end[-1] == L'|')
        end--;

    cell_start = start;
    p = start;

    while (p <= end &&
           count < max_cells) {

        if (p == end || *p == L'|') {
            cells[count].start = cell_start;
            cells[count].end = p;

            trim_cell(&cells[count]);

            count++;
            cell_start = p + 1;
        }

        p++;
    }

    return count;
}

static BOOL is_table_separator(
    const WCHAR *start,
    const WCHAR *end,
    int expected_cols)
{
    CELL cells[MAX_COLS];
    int cols;

    if (!line_has_pipe(start, end))
        return FALSE;

    cols = split_cells(
        start,
        end,
        cells,
        MAX_COLS
    );

    if (cols < 2)
        return FALSE;

    if (expected_cols > 0 &&
        cols != expected_cols) {
        return FALSE;
    }

    for (int i = 0; i < cols; i++) {
        const WCHAR *p = cells[i].start;
        int dashes = 0;

        if (p == cells[i].end)
            return FALSE;

        while (p < cells[i].end) {
            if (*p == L'-') {
                dashes++;
            }
            else if (*p == L':' ||
                     *p == L' ' ||
                     *p == L'\t') {
                /* allowed */
            }
            else {
                return FALSE;
            }

            p++;
        }

        if (dashes < 3)
            return FALSE;
    }

    return TRUE;
}


/* -------------------------------------------------------------
   Native RichEdit table insertion
   ------------------------------------------------------------- */

typedef struct {
    char *data;
    size_t len;
    size_t cap;
} RTFBUF;

typedef struct {
    const char *data;
    size_t len;
    size_t pos;
} RTFSTREAM;

static BOOL rtf_reserve(RTFBUF *b, size_t extra)
{
    size_t need = b->len + extra + 1;
    size_t cap;
    char *p;

    if (need <= b->cap)
        return TRUE;

    cap = b->cap ? b->cap : 1024;
    while (cap < need)
        cap *= 2;

    p = (char *)realloc(b->data, cap);
    if (!p)
        return FALSE;

    b->data = p;
    b->cap = cap;
    return TRUE;
}

static BOOL rtf_puts(RTFBUF *b, const char *s)
{
    size_t n = strlen(s);

    if (!rtf_reserve(b, n))
        return FALSE;

    memcpy(b->data + b->len, s, n);
    b->len += n;
    b->data[b->len] = 0;
    return TRUE;
}

static BOOL rtf_put_int(RTFBUF *b, int v)
{
    char tmp[64];
    int n = snprintf(tmp, sizeof(tmp), "%d", v);

    if (n < 0)
        return FALSE;

    return rtf_puts(b, tmp);
}

static BOOL rtf_put_char(RTFBUF *b, WCHAR wc)
{
    if (wc == L'\\' || wc == L'{' || wc == L'}') {
        char tmp[3];
        tmp[0] = '\\';
        tmp[1] = (char)wc;
        tmp[2] = 0;
        return rtf_puts(b, tmp);
    }

    if (wc >= 0x20 && wc <= 0x7e) {
        char tmp[2];
        tmp[0] = (char)wc;
        tmp[1] = 0;
        return rtf_puts(b, tmp);
    }

    {
        char tmp[64];
        short sw = (short)wc;
        int n = snprintf(tmp, sizeof(tmp), "\\u%d?", (int)sw);

        if (n < 0)
            return FALSE;

        return rtf_puts(b, tmp);
    }
}

static BOOL rtf_put_text(
    RTFBUF *b,
    const WCHAR *start,
    const WCHAR *end)
{
    while (start < end) {
        if (!rtf_put_char(b, *start))
            return FALSE;
        start++;
    }

    return TRUE;
}

static BOOL rtf_put_inline_markdown(
    RTFBUF *b,
    const WCHAR *start,
    const WCHAR *end,
    BOOL base_bold)
{
    const WCHAR *p = start;
    const WCHAR *chunk = start;
    BOOL bold = base_bold;
    BOOL italic = FALSE;
    BOOL strike = FALSE;
    BOOL code = FALSE;

    if (bold && !rtf_puts(b, "\\b "))
        return FALSE;

    while (p < end) {
        int marker_len = 0;
        int mark = 0;

        if (!code) {
            if (p + 2 < end &&
                ((p[0] == L'*' && p[1] == L'*' && p[2] == L'*') ||
                 (p[0] == L'_' && p[1] == L'_' && p[2] == L'_'))) {
                mark = 3;
                marker_len = 3;
            }
            else if (p + 1 < end &&
                     ((p[0] == L'*' && p[1] == L'*') ||
                      (p[0] == L'_' && p[1] == L'_'))) {
                mark = 2;
                marker_len = 2;
            }
            else if (p + 1 < end && p[0] == L'~' && p[1] == L'~') {
                mark = 4;
                marker_len = 2;
            }
            else if (*p == L'*' || *p == L'_') {
                mark = 1;
                marker_len = 1;
            }
            else if (*p == L'`') {
                mark = 5;
                marker_len = 1;
            }
        }
        else if (*p == L'`') {
            mark = 5;
            marker_len = 1;
        }

        if (mark) {
            if (p > chunk && !rtf_put_text(b, chunk, p))
                return FALSE;

            if (mark == 1) {
                italic = !italic;
                if (!rtf_puts(b, italic ? "\\i " : "\\i0 "))
                    return FALSE;
            }
            else if (mark == 2) {
                bold = !bold;
                if (!rtf_puts(b, bold ? "\\b " : "\\b0 "))
                    return FALSE;
            }
            else if (mark == 3) {
                bold = !bold;
                italic = !italic;
                if (!rtf_puts(b, bold ? "\\b " : "\\b0 "))
                    return FALSE;
                if (!rtf_puts(b, italic ? "\\i " : "\\i0 "))
                    return FALSE;
            }
            else if (mark == 4) {
                strike = !strike;
                if (!rtf_puts(b, strike ? "\\strike " : "\\strike0 "))
                    return FALSE;
            }
            else if (mark == 5) {
                code = !code;
                if (!rtf_puts(b, code ? "\\f1\\highlight1 " : "\\f0\\highlight0 "))
                    return FALSE;
            }

            p += marker_len;
            chunk = p;
            continue;
        }

        p++;
    }

    if (p > chunk && !rtf_put_text(b, chunk, p))
        return FALSE;

    return rtf_puts(b, "\\b0\\i0\\strike0\\f0\\highlight0 ");
}

static DWORD CALLBACK rtf_stream_in_cb(
    DWORD_PTR cookie,
    LPBYTE pbBuff,
    LONG cb,
    LONG *pcb)
{
    RTFSTREAM *s = (RTFSTREAM *)cookie;
    size_t remain = s->len - s->pos;
    size_t n = remain < (size_t)cb ? remain : (size_t)cb;

    if (n) {
        memcpy(pbBuff, s->data + s->pos, n);
        s->pos += n;
    }

    *pcb = (LONG)n;
    return 0;
}

static BOOL stream_rtf_selection(
    HWND edit,
    const char *data,
    size_t len)
{
    EDITSTREAM es;
    RTFSTREAM st;
    LRESULT result;

    st.data = data;
    st.len = len;
    st.pos = 0;

    ZeroMemory(&es, sizeof(es));
    es.dwCookie = (DWORD_PTR)&st;
    es.pfnCallback = rtf_stream_in_cb;

    select_end(edit);

    result = SendMessageW(
        edit,
        EM_STREAMIN,
        SF_RTF | SFF_SELECTION,
        (LPARAM)&es
    );

    return es.dwError == 0 && result != 0;
}

static BOOL insert_native_table(
    HWND edit,
    const WCHAR *header,
    const WCHAR *body,
    const WCHAR *table_end,
    int cols,
    int rows)
{
    RTFBUF b = {0};
    RECT rc;
    HDC dc;
    int dpi;
    int width_px;
    int width_twips;
    int cell_twips;
    const WCHAR *line = header;

    if (cols < 1 || rows < 1)
        return FALSE;

    GetClientRect(edit, &rc);
    width_px = rc.right - rc.left - INNER_MARGIN * 2 - 12;

    if (width_px < cols * 40)
        width_px = cols * 40;

    dc = GetDC(edit);
    dpi = GetDeviceCaps(dc, LOGPIXELSX);
    ReleaseDC(edit, dc);

    width_twips = MulDiv(width_px, 1440, dpi);
    cell_twips = width_twips / cols;

    if (!rtf_puts(&b,
        "{\\rtf1\\ansi\\deff0\\uc1"
        "{\\fonttbl{\\f0\\fcharset0 Arial;}{\\f1\\fcharset0 Consolas;}}"
        "{\\colortbl;\\red240\\green240\\blue240;}"))
        goto fail;

    for (int r = 0; r < rows; r++) {
        const WCHAR *end = line_end(line);
        CELL cells[MAX_COLS];
        int n = split_cells(line, end, cells, MAX_COLS);

        if (n != cols)
            goto fail;

        if (!rtf_puts(&b,
            "\\trowd\\trgaph80\\trleft0"
            "\\trpaddl80\\trpaddr80\\trpaddt40\\trpaddb40"
            "\\trpaddfl3\\trpaddfr3\\trpaddft3\\trpaddfb3 "))
            goto fail;

        for (int c = 0; c < cols; c++) {
            int cellx = cell_twips * (c + 1);

            if (!rtf_puts(&b,
                "\\clvertalt"
                "\\clbrdrt\\brdrs\\brdrw10"
                "\\clbrdrl\\brdrs\\brdrw10"
                "\\clbrdrb\\brdrs\\brdrw10"
                "\\clbrdrr\\brdrs\\brdrw10"
                "\\cellx"))
                goto fail;

            if (!rtf_put_int(&b, cellx) || !rtf_puts(&b, " "))
                goto fail;
        }

        if (!rtf_puts(&b, "\\pard\\intbl\\f0\\fs20 "))
            goto fail;

        for (int c = 0; c < cols; c++) {
            if (!rtf_put_inline_markdown(
                    &b,
                    cells[c].start,
                    cells[c].end,
                    r == 0))
                goto fail;

            if (!rtf_puts(&b, "\\cell "))
                goto fail;
        }

        if (!rtf_puts(&b, "\\row "))
            goto fail;

        if (r == 0)
            line = body;
        else
            line = next_line(line);
    }

    if (!rtf_puts(&b, "\\pard\\f0\\fs20\\par}"))
        goto fail;

    if (!stream_rtf_selection(edit, b.data, b.len))
        goto fail;

    free(b.data);
    return TRUE;

fail:
    free(b.data);
    return FALSE;
}

static const WCHAR *render_native_table(
    HWND edit,
    const WCHAR *p)
{
    const WCHAR *header_end =
        line_end(p);

    CELL header_cells[MAX_COLS];

    int cols =
        split_cells(
            p,
            header_end,
            header_cells,
            MAX_COLS
        );

    if (cols < 2 ||
        !line_has_pipe(
            p,
            header_end)) {
        return NULL;
    }

    const WCHAR *separator =
        next_line(p);

    if (!*separator)
        return NULL;

    const WCHAR *separator_end =
        line_end(separator);

    if (!is_table_separator(
            separator,
            separator_end,
            cols)) {
        return NULL;
    }

    const WCHAR *body =
        next_line(separator);

    const WCHAR *scan = body;
    int body_rows = 0;

    while (*scan) {
        const WCHAR *end = line_end(scan);
        CELL cells[MAX_COLS];

        if (!line_has_pipe(
                scan,
                end)) {
            break;
        }

        int n =
            split_cells(
                scan,
                end,
                cells,
                MAX_COLS
            );

        if (n != cols)
            break;

        body_rows++;
        scan = next_line(scan);
    }

    const WCHAR *table_end = scan;
    int rows = 1 + body_rows;

    if (!insert_native_table(
            edit,
            p,
            body,
            table_end,
            cols,
            rows)) {
        return NULL;
    }

    return table_end;
}


/* -------------------------------------------------------------
   Markdown renderer
   ------------------------------------------------------------- */

static void render_markdown(
    HWND edit,
    const WCHAR *text)
{
    const WCHAR *p = text;

    SendMessageW(
        edit,
        EM_SETREADONLY,
        FALSE,
        0
    );

    SetWindowTextW(
        edit,
        L""
    );

    SendMessageW(
        edit,
        WM_SETREDRAW,
        FALSE,
        0
    );

    while (*p) {
        const WCHAR *after_table =
            render_native_table(
                edit,
                p
            );

        if (after_table) {
            p = after_table;
            continue;
        }

        {
            const WCHAR *start = p;
            const WCHAR *end = line_end(p);
            const WCHAR *content = start;
            const WCHAR *list_content;
            int list_number = 0;
            int level = 0;

            if (is_blank_line(
                    start,
                    end)) {

                append_normal_line(
                    edit,
                    start,
                    end,
                    10,
                    FALSE
                );

                p = next_line(p);
                continue;
            }

            while (content < end &&
                   *content == L'#' &&
                   level < 4) {

                level++;
                content++;
            }

            if (level > 0 &&
                content < end &&
                (*content == L' ' ||
                 *content == L'\t')) {

                while (content < end &&
                       (*content == L' ' ||
                        *content == L'\t')) {
                    content++;
                }

                append_heading_line(
                    edit,
                    content,
                    end,
                    level
                );

                p = next_line(p);
                continue;
            }

            {
                const WCHAR *q = start;

                while (q < end &&
                       (*q == L' ' ||
                        *q == L'\t')) {
                    q++;
                }

                if (q < end && *q == L'>') {
                    append_quote_line(
                        edit,
                        start,
                        end
                    );

                    p = next_line(p);
                    continue;
                }
            }

            if (parse_unordered_list(
                    start,
                    end,
                    &list_content)) {

                append_list_line(
                    edit,
                    list_content,
                    end,
                    FALSE,
                    0
                );

                p = next_line(p);
                continue;
            }

            if (parse_ordered_list(
                    start,
                    end,
                    &list_content,
                    &list_number)) {

                append_list_line(
                    edit,
                    list_content,
                    end,
                    TRUE,
                    list_number
                );

                p = next_line(p);
                continue;
            }

            if (is_horizontal_rule(
                    start,
                    end)) {

                append_horizontal_rule(edit);

                p = next_line(p);
                continue;
            }

            append_normal_line(
                edit,
                start,
                end,
                10,
                FALSE
            );
        }

        p = next_line(p);
    }

    SendMessageW(
        edit,
        EM_SETSEL,
        0,
        0
    );

    SendMessageW(
        edit,
        WM_SETREDRAW,
        TRUE,
        0
    );

    SendMessageW(
        edit,
        EM_SETREADONLY,
        TRUE,
        0
    );

    InvalidateRect(
        edit,
        NULL,
        TRUE
    );
}


/* -------------------------------------------------------------
   File loading
   ------------------------------------------------------------- */

static WCHAR *load_text_file(
    const WCHAR *filename)
{
    HANDLE file;
    LARGE_INTEGER size;
    BYTE *data = NULL;
    WCHAR *text = NULL;
    DWORD bytes_read;
    DWORD length;
    BYTE *src;
    DWORD src_len;
    UINT codepage = CP_UTF8;
    DWORD flags = MB_ERR_INVALID_CHARS;
    int chars;

    file = CreateFileW(
        filename,
        GENERIC_READ,
        FILE_SHARE_READ,
        NULL,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        NULL
    );

    if (file == INVALID_HANDLE_VALUE)
        return NULL;

    if (!GetFileSizeEx(file, &size) ||
        size.QuadPart > 0x7fffffff) {

        CloseHandle(file);
        return NULL;
    }

    length =
        (DWORD)size.QuadPart;

    data =
        (BYTE *)malloc(
            length + 1
        );

    if (!data) {
        CloseHandle(file);
        return NULL;
    }

    if (!ReadFile(
            file,
            data,
            length,
            &bytes_read,
            NULL)) {

        free(data);
        CloseHandle(file);

        return NULL;
    }

    CloseHandle(file);

    data[bytes_read] = 0;

    src = data;
    src_len = bytes_read;

    if (src_len >= 3 &&
        src[0] == 0xEF &&
        src[1] == 0xBB &&
        src[2] == 0xBF) {

        src += 3;
        src_len -= 3;
    }

    chars =
        MultiByteToWideChar(
            codepage,
            flags,
            (LPCCH)src,
            src_len,
            NULL,
            0
        );

    if (!chars) {
        codepage = CP_ACP;
        flags = 0;

        chars =
            MultiByteToWideChar(
                codepage,
                flags,
                (LPCCH)src,
                src_len,
                NULL,
                0
            );
    }

    if (!chars) {
        free(data);
        return NULL;
    }

    text =
        (WCHAR *)malloc(
            (chars + 1) *
            sizeof(WCHAR)
        );

    if (!text) {
        free(data);
        return NULL;
    }

    MultiByteToWideChar(
        codepage,
        flags,
        (LPCCH)src,
        src_len,
        text,
        chars
    );

    text[chars] = 0;

    free(data);

    return text;
}



/* -------------------------------------------------------------
   Save current RichEdit contents as RTF
   ------------------------------------------------------------- */

static DWORD CALLBACK stream_out_callback(
    DWORD_PTR cookie,
    LPBYTE buffer,
    LONG cb,
    LONG *pcb)
{
    HANDLE file = (HANDLE)cookie;
    DWORD written = 0;

    if (!WriteFile(
            file,
            buffer,
            (DWORD)cb,
            &written,
            NULL)) {

        *pcb = 0;
        return 1;
    }

    *pcb = (LONG)written;
    return written == (DWORD)cb ? 0 : 1;
}

static BOOL save_richedit_as_rtf(
    HWND edit,
    const WCHAR *filename)
{
    HANDLE file;
    EDITSTREAM es;
    LRESULT result;

    file = CreateFileW(
        filename,
        GENERIC_WRITE,
        0,
        NULL,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        NULL
    );

    if (file == INVALID_HANDLE_VALUE)
        return FALSE;

    ZeroMemory(&es, sizeof(es));
    es.dwCookie = (DWORD_PTR)file;
    es.pfnCallback = stream_out_callback;

    result = SendMessageW(
        edit,
        EM_STREAMOUT,
        SF_RTF,
        (LPARAM)&es
    );

    CloseHandle(file);

    return es.dwError == 0 && result != 0;
}

static void save_rtf(HWND hwnd)
{
    WCHAR filename[MAX_PATH] = L"document.rtf";
    OPENFILENAMEW ofn;

    ZeroMemory(&ofn, sizeof(ofn));

    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = filename;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter =
        L"Rich Text Format (*.rtf)\0*.rtf\0"
        L"All files (*.*)\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrDefExt = L"rtf";
    ofn.Flags =
        OFN_OVERWRITEPROMPT |
        OFN_PATHMUSTEXIST |
        OFN_HIDEREADONLY;

    if (!GetSaveFileNameW(&ofn))
        return;

    if (!save_richedit_as_rtf(g_edit, filename)) {
        MessageBoxW(
            hwnd,
            L"Cannot save RTF file.",
            L"MD Viewer",
            MB_OK | MB_ICONERROR
        );
    }
}

/* -------------------------------------------------------------
   Open file
   ------------------------------------------------------------- */

static void open_file(HWND hwnd)
{
    WCHAR filename[MAX_PATH] = L"";
    OPENFILENAMEW ofn;
    WCHAR *text;
    const WCHAR *name;
    WCHAR title[MAX_PATH + 32];

    ZeroMemory(
        &ofn,
        sizeof(ofn)
    );

    ofn.lStructSize =
        sizeof(ofn);

    ofn.hwndOwner =
        hwnd;

    ofn.lpstrFile =
        filename;

    ofn.nMaxFile =
        MAX_PATH;

    ofn.lpstrFilter =
        L"Markdown files (*.md)\0*.md\0"
        L"Text files (*.txt)\0*.txt\0"
        L"All files (*.*)\0*.*\0";

    ofn.nFilterIndex = 1;

    ofn.Flags =
        OFN_FILEMUSTEXIST |
        OFN_PATHMUSTEXIST |
        OFN_HIDEREADONLY;

    if (!GetOpenFileNameW(&ofn))
        return;

    text =
        load_text_file(
            filename
        );

    if (!text) {
        MessageBoxW(
            hwnd,
            L"Cannot read file.",
            L"MD Viewer",
            MB_OK |
            MB_ICONERROR
        );

        return;
    }

    render_markdown(
        g_edit,
        text
    );

    free(text);

    name =
        wcsrchr(
            filename,
            L'\\'
        );

    if (name)
        name++;
    else
        name = filename;

    swprintf(
        title,
        MAX_PATH + 32,
        L"%ls - MD Viewer",
        name
    );

    SetWindowTextW(
        hwnd,
        title
    );
}


/* -------------------------------------------------------------
   Window procedure
   ------------------------------------------------------------- */

static LRESULT CALLBACK wndproc(
    HWND hwnd,
    UINT msg,
    WPARAM wParam,
    LPARAM lParam)
{
    switch (msg) {

    case WM_CREATE:

        g_edit =
            CreateWindowExW(
                0,
                MSFTEDIT_CLASS,
                L"",
                WS_CHILD |
                WS_VISIBLE |
                WS_VSCROLL |
                ES_MULTILINE |
                ES_AUTOVSCROLL |
                ES_NOHIDESEL,
                0,
                0,
                0,
                0,
                hwnd,
                NULL,
                GetModuleHandleW(NULL),
                NULL
            );

        if (!g_edit)
            return -1;

        SendMessageW(
            g_edit,
            EM_SETMARGINS,
            EC_LEFTMARGIN |
            EC_RIGHTMARGIN,
            MAKELPARAM(
                INNER_MARGIN,
                INNER_MARGIN
            )
        );

        SendMessageW(
            g_edit,
            EM_SETBKGNDCOLOR,
            0,
            RGB(255, 255, 255)
        );

        SendMessageW(
            g_edit,
            EM_SETREADONLY,
            TRUE,
            0
        );

        return 0;


    case WM_SIZE:
    {
        int width =
            LOWORD(lParam);

        int height =
            HIWORD(lParam);

        int edit_width =
            width -
            OUTER_MARGIN * 2;

        int edit_height =
            height -
            OUTER_MARGIN * 2;

        if (edit_width < 0)
            edit_width = 0;

        if (edit_height < 0)
            edit_height = 0;

        MoveWindow(
            g_edit,
            OUTER_MARGIN,
            OUTER_MARGIN,
            edit_width,
            edit_height,
            TRUE
        );

        return 0;
    }


    case WM_COMMAND:

        if (LOWORD(wParam) == ID_OPEN) {
            open_file(hwnd);
            return 0;
        }

        if (LOWORD(wParam) == ID_SAVE_RTF) {
            save_rtf(hwnd);
            return 0;
        }

        break;


    case WM_DESTROY:

        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(
        hwnd,
        msg,
        wParam,
        lParam
    );
}


/* -------------------------------------------------------------
   Entry point
   ------------------------------------------------------------- */

int WINAPI wWinMain(
    HINSTANCE instance,
    HINSTANCE prev,
    PWSTR cmdline,
    int show)
{
    const WCHAR CLASS_NAME[] =
        L"MinimalMDViewer";

    WNDCLASSW wc;
    HMENU menu;
    HMENU filemenu;
    HWND hwnd;
    MSG msg;

    (void)prev;
    (void)cmdline;

    g_richedit =
        LoadLibraryW(
            L"Msftedit.dll"
        );

    if (!g_richedit) {
        MessageBoxW(
            NULL,
            L"Cannot load Msftedit.dll.",
            L"MD Viewer",
            MB_OK |
            MB_ICONERROR
        );

        return 1;
    }

    ZeroMemory(
        &wc,
        sizeof(wc)
    );

    wc.lpfnWndProc =
        wndproc;

    wc.hInstance =
        instance;

    wc.lpszClassName =
        CLASS_NAME;

    wc.hCursor =
        LoadCursorW(
            NULL,
            IDC_ARROW
        );

    wc.hbrBackground =
        (HBRUSH)(
            COLOR_WINDOW + 1
        );

    if (!RegisterClassW(&wc)) {
        FreeLibrary(g_richedit);
        return 1;
    }

    menu =
        CreateMenu();

    filemenu =
        CreatePopupMenu();

    AppendMenuW(
        filemenu,
        MF_STRING,
        ID_OPEN,
        L"&Open..."
    );

    AppendMenuW(
        filemenu,
        MF_STRING,
        ID_SAVE_RTF,
        L"&Save as RTF..."
    );

    AppendMenuW(
        menu,
        MF_POPUP,
        (UINT_PTR)filemenu,
        L"&File"
    );

    hwnd =
        CreateWindowExW(
            0,
            CLASS_NAME,
            L"MD Viewer",
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            800,
            600,
            NULL,
            menu,
            instance,
            NULL
        );

    if (!hwnd) {
        FreeLibrary(
            g_richedit
        );

        return 1;
    }

    ShowWindow(
        hwnd,
        show
    );

    UpdateWindow(
        hwnd
    );

    while (
        GetMessageW(
            &msg,
            NULL,
            0,
            0
        ) > 0) {

        TranslateMessage(
            &msg
        );

        DispatchMessageW(
            &msg
        );
    }

    FreeLibrary(
        g_richedit
    );

    return (int)msg.wParam;
}
