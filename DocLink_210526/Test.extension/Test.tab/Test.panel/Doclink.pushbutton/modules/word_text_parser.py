# -*- coding: utf-8 -*-
import os
import _imports
from logger import LogManager

# Late binding for Word
_WORD_AVAILABLE = _imports._WORD_AVAILABLE

def parse_word(word_path, selected_page=None):
    """
    Parse a Word document into a GridModel-compatible dict.
    Preserves order of text and tables.
    If selected_page is an integer, only returns content from that page.
    """
    if not _WORD_AVAILABLE:
        raise RuntimeError("Word COM not available.")

    from _imports import Word, Marshal
    
    word_app = None
    doc = None
    try:
        LogManager.debug("[WordParser] Opening document: {}".format(os.path.basename(word_path)))
        word_app = Word.ApplicationClass()
        word_app.Visible = False
        doc = word_app.Documents.Open(word_path, False, True)
        
        all_cells = []
        max_cols = 1
        current_row = 0
        col_widths_ft = []

        # We'll use table start position as a unique ID
        processed_table_starts = set()

        para_count = doc.Paragraphs.Count
        LogManager.debug("[WordParser] Total paragraphs to process: {}".format(para_count))

        # Iterate through all paragraphs in the main story
        for i in range(1, para_count + 1):
            if i % 50 == 0:
                LogManager.info("Processing Word content: {}%".format(int((i/float(para_count))*100)))

            try:
                para = doc.Paragraphs.Item(i)
                r = para.Range
                
                # Check for page filtering
                if selected_page:
                    # wdActiveEndPageNumber = 3
                    curr_p = r.Information(3)
                    if int(curr_p) != int(selected_page):
                        continue

                # wdWithInTable = 12
                if r.Information(12): 
                    # This paragraph is inside a table
                    table = r.Tables.Item(1)
                    t_start = table.Range.Start
                    
                    if t_start not in processed_table_starts:
                        processed_table_starts.add(t_start)
                        
                        # Process the whole table
                        t_cols = 0
                        try:
                            t_cols = table.Columns.Count
                        except Exception:
                            # Fallback if merger cells exist
                            for r_test in range(1, min(table.Rows.Count + 1, 5)):
                                try:
                                    c_count = table.Rows.Item(r_test).Cells.Count
                                    if c_count > t_cols: t_cols = c_count
                                except Exception: pass
                        
                        if t_cols <= 0: t_cols = 1
                        if t_cols > max_cols: max_cols = t_cols
                        
                        # Column widths
                        t_widths = []
                        try:
                            ref_row = table.Rows.Item(1)
                            for c_ref in range(1, ref_row.Cells.Count + 1):
                                w_pt = ref_row.Cells.Item(c_ref).Width
                                t_widths.append(w_pt / 72.0 / 12.0)
                        except Exception:
                            t_widths = [1.0 / 12.0] * t_cols
                        
                        if len(t_widths) > len(col_widths_ft):
                            col_widths_ft = t_widths
                        
                        for r_idx in range(1, table.Rows.Count + 1):
                            try:
                                row = table.Rows.Item(r_idx)
                                for c_idx in range(1, row.Cells.Count + 1):
                                    cell = row.Cells.Item(c_idx)
                                    # Support bullets in cells
                                    prefix = cell.Range.ListFormat.ListString
                                    txt = cell.Range.Text.strip('\r\x07')
                                    if prefix: txt = prefix + " " + txt
                                    
                                    h_al = {0:'left', 1:'center', 2:'right', 3:'justify'}.get(int(cell.Range.ParagraphFormat.Alignment), 'left')
                                    v_al = {0:'top', 1:'center', 2:'bottom'}.get(int(cell.VerticalAlignment), 'top')

                                    all_cells.append({
                                        'row': current_row,
                                        'col': c_idx - 1,
                                        'value': txt,
                                        'font_bold': bool(cell.Range.Font.Bold),
                                        'font_italic': bool(cell.Range.Font.Italic),
                                        'font_underline': bool(cell.Range.Font.Underline),
                                        'font_size': float(cell.Range.Font.Size),
                                        'font_name': str(cell.Range.Font.Name),
                                        'h_align': h_al,
                                        'v_align': v_al,
                                        'is_master': True
                                    })
                                current_row += 1
                            except Exception:
                                continue
                else:
                    # Normal paragraph
                    prefix = para.Range.ListFormat.ListString
                    txt = para.Range.Text.strip('\r\n')
                    if prefix: 
                        txt = prefix + " " + txt
                    
                    if not txt and all_cells and not all_cells[-1]['value']:
                        continue
                    
                    h_al = {0:'left', 1:'center', 2:'right', 3:'justify'}.get(int(para.Range.ParagraphFormat.Alignment), 'left')

                    all_cells.append({
                        'row': current_row,
                        'col': 0,
                        'value': txt,
                        'font_bold': bool(para.Range.Font.Bold),
                        'font_italic': bool(para.Range.Font.Italic),
                        'font_underline': bool(para.Range.Font.Underline),
                        'font_size': float(para.Range.Font.Size),
                        'font_name': str(para.Range.Font.Name),
                        'h_align': h_al,
                        'v_align': 'top',
                        'is_master': True
                    })
                    current_row += 1
            except Exception as ex:
                LogManager.debug("[WordParser] Skipping paragraph {}: {}".format(i, ex))
                continue

        if not col_widths_ft:
            col_widths_ft = [5.0 / 12.0] * max_cols
        elif len(col_widths_ft) < max_cols:
            col_widths_ft.extend([col_widths_ft[-1]] * (max_cols - len(col_widths_ft)))

        LogManager.debug("[WordParser] Extraction complete. Rows: {}, Cols: {}".format(current_row, max_cols))

        return {
            'rows': current_row,
            'cols': max_cols,
            'cells': all_cells,
            'col_widths_ft': col_widths_ft,
            'row_heights_ft': [20.0 / 72.0 / 12.0] * current_row,
            'default_font_name': 'Calibri',
            'default_font_size': 11.0,
            'warnings': []
        }

    finally:
        try:
            if doc: doc.Close(False)
        except Exception:
            pass
        try:
            if word_app:
                try:
                    word_app.Quit()
                except Exception:
                    pass
                Marshal.ReleaseComObject(word_app)
            LogManager.debug("[WordParser] Word application closed.")
        except Exception:
            pass

def get_word_info(word_path):
    """
    Returns basic info about a Word doc (e.g. Page Count).
    """
    if not _WORD_AVAILABLE: return {'pages': 0}
    from _imports import Word
    word_app = None
    doc = None
    try:
        word_app = Word.ApplicationClass()
        word_app.Visible = False
        doc = word_app.Documents.Open(word_path, False, True)
        # wdStatisticPages = 2
        pages = doc.ComputeStatistics(2)
        return {'pages': int(pages)}
    except Exception:
        return {'pages': 0}
    finally:
        try:
            if doc: doc.Close(False)
        except Exception:
            pass
        try:
            if word_app:
                try:
                    word_app.Quit()
                except Exception:
                    pass
                from _imports import Marshal
                Marshal.ReleaseComObject(word_app)
        except Exception:
            pass
