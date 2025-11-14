import subprocess
import json
import difflib
import re 
import os
import sys
def resource_path(relative_path):
    """
    Get absolute path to resource.
    Works for both development (Python) and PyInstaller-built .exe
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class SmartLLMEngine:
    def __init__(self, model_name="gemma:2b", label_list_path="label_mapping.txt"):
        self.model_name = model_name
        label_list_path = resource_path(label_list_path)

        # Load labels from file
        with open(label_list_path, "r", encoding="utf-8") as f:
            self.labels = [line.strip().lstrip("- ").strip() for line in f if line.strip()]

        # Manual semantic hints (map label -> list of keywords/synonyms)
        self.semantic_hints = {
            "add_ins": ["add-ins", "addin", "addin tab", "extensions"],
    "design_effect": ["design effect", "style effect"],
    "design_styles": ["design style", "themes", "style set","document styles"],
    "font_family_dropdown": ["font family", "change font", "font style", "font type","change text"],
    "font_size_dropdown": ["font size", "increase font", "decrease font", "text size"],
    "icon_3dmodels": ["3D models", "3d object", "insert 3d", "3d shapes"],
    "icon_add_text": ["add text","type text"],
    "icon_align": ["align text", "justify text", "center text", "left align", "right align"],
    "icon_arrange_all": ["arrange all", "tile windows", "window arrangement"],
    "icon_bibliography": ["bibliography", "insert citation", "works cited"],
    "icon_black_page": ["black page", "insert blank page", "empty page", "new page"],
    "icon_bold": ["bold", "make text bold", "strong text","bold text"],
    "icon_bookmarks": ["bookmark", "add bookmark", "mark location", "bookmark text"],
    "icon_borders": ["borders", "add border", "cell border"],
    "icon_break_page": ["page break", "break pages", "insert page break","separate pages", "new page"],
    "icon_breaks": ["line break", "section break", "breaks"],
    "icon_bullets": ["bullets", "bullet list", "unordered list", "dot list"],
    "icon_change_accept": ["accept change", "track changes accept", "accept editing"],
    "icon_change_case": ["change case", "uppercase", "lowercase", "capitalize"],
    "icon_change_next": ["next change", "next edit", "track change next"],
    "icon_change_previous": ["previous change", "previous edit", "track change previous"],
    "icon_change_provider": ["change provider", "track provider change"],
    "icon_change_reject": ["reject change", "track changes reject", "reject edit"],
    "icon_chart": ["chart", "insert chart", "graph", "diagram"],
    "icon_check_accessibility": ["check accessibility", "accessibility check", "accessible"],
    "icon_clear_format": ["clear formatting", "remove format", "reset style"],
    "icon_colors": ["color", "highlight color"],
    "icon_columns": ["columns", "split text into columns", "multi-column"],
    "icon_comments": ["comments", "add comment", "review comment"],
    "icon_compare": ["compare document", "track differences", "compare changes"],
    "icon_contact_support": ["contact support", "help", "support"],
    "icon_copy": ["copy", "duplicate text", "copy selection"],
    "icon_cover_page": ["cover page", "insert cover page", "title page"],
    "icon_cross_reference": ["cross reference", "insert reference", "reference link"],
    "icon_cut": ["cut", "remove selection", "delete and copy"],
    "icon_date_time": ["insert date", "insert time", "date and time"],
    "icon_draft_view": ["draft view", "view draft", "simple layout"],
    "icon_draw_eraser": ["eraser", "delete drawing", "remove ink"],
    "icon_draw_select": ["draw select", "select drawing", "select ink"],
    "icon_drawing_canvas": ["drawing canvas", "canvas", "drawing area"],
    "icon_drop_cap": ["drop cap", "large initial", "first letter large"],
    "icon_dropdown_style": ["dropdown style", "style dropdown", "format style"],
    "icon_equation": ["equation", "insert equation", "math formula"],
    "icon_feedback": ["feedback", "send feedback", "report"],
    "icon_filter_markup": ["filter markup", "show changes", "track changes filter"],
    "icon_find": ["find", "search", "search text"],
    "icon_focus_mode": ["focus mode", "distraction free", "reading mode"],
    "icon_footer": ["footer", "insert footer", "page footer"],
    "icon_format_background": ["format background", "page background", "change background"],
    "icon_format_painter": ["format painter", "copy formatting", "paint format"],
    "icon_forward_backward": ["bring forward", "send backward", "layer order"],
    "icon_get_word_mobile_app": ["word mobile", "mobile app", "get word app"],
    "icon_gridlines": ["gridlines", "show gridlines", "view grid"],
    "icon_group": ["group", "group objects", "combine shapes"],
    "icon_header": ["header", "insert header", "page header"],
    "icon_hide_ink": ["hide ink", "hide drawing", "hide annotations"],
    "icon_highlight": ["highlight", "text highlight", "highlight color"],
    "icon_hyphenation": ["hyphenation", "automatic hyphen", "split words"],
    "icon_icons": ["icons", "insert icons", "icon gallery"],
    "icon_immersive_reader": ["immersive reader", "read aloud", "learning mode"],
    "icon_indent": ["indent", "increase indent", "decrease indent"],
    "icon_ink_help": ["ink help", "drawing help", "ink assistance"],
    "icon_ink_replay": ["ink replay", "play drawing", "replay ink"],
    "icon_ink_to_math": ["ink to math", "convert drawing to equation"],
    "icon_ink_to_shape": ["ink to shape", "convert drawing to shape"],
    "icon_insert_caption": ["caption", "insert caption", "figure caption"],
    "icon_insert_citation": ["citation", "insert citation", "reference"],
    "icon_insert_endnote": ["endnote", "insert endnote", "footnote reference"],
    "icon_insert_footnote": ["footnote", "insert footnote", "reference note"],
    "icon_insert_index": ["index", "insert index", "table of index"],
    "icon_insert_table_figures": ["table of figures", "insert table of figures"],
    "icon_italic": ["italic", "italicize text", "slant text"],
    "icon_language": ["language", "change language", "proofing language"],
    "icon_line_number": ["line number", "number lines", "show line numbers"],
    "icon_line_spacing": ["line spacing", "paragraph spacing", "line height"],
    "icon_links": ["hyperlink", "insert link", "add link"],
    "icon_macros": ["macros", "create macro", "run macro"],
    "icon_manage_source": ["manage source", "citation source", "reference source"],
    "icon_margins": ["margins", "page margins", "set margins"],
    "icon_mark_citation": ["mark citation", "tag citation", "reference marker"],
    "icon_mark_entry": ["mark entry", "index entry", "mark index"],
    "icon_microsoft_help": ["microsoft help", "help", "word help"],
    "icon_navigation_pane": ["navigation pane", "document map", "view pane"],
    "icon_new_window": ["new window", "open new window", "duplicate window"],
    "icon_next_footnote": ["next footnote", "jump to footnote", "footnote next"],
    "icon_objects": ["objects", "insert object", "embedded object"],
    "icon_online_videos": ["online videos", "insert video", "stream video"],
    "icon_orientation": ["orientation", "page orientation", "portrait", "landscape"],
    "icon_outline_view": ["outline view", "document outline", "view structure"],
    "icon_page_border": ["page border", "border", "set page border"],
    "icon_page_color": ["page color", "background color", "change page color"],
    "icon_page_movement": ["page movement", "scroll pages", "move pages"],
    "icon_page_number": ["page number", "insert page number", "number pages"],
    "icon_page_size": ["page size", "paper size", "set page size"],
    "icon_page_width": ["page width", "adjust page width", "resize page width"],
    "icon_paragraph_spacing": ["paragraph spacing", "space between paragraphs", "line spacing"],
    "icon_paste": ["paste", "paste clipboard", "insert copied text"],
    "icon_pens": ["pens", "drawing pens", "ink pen"],
    "icon_pictures": ["insert pictures", "add image", "insert image","picture"],
    "icon_position": ["position", "object position", "arrange object"],
    "icon_print_layout": ["print layout", "page layout", "layout view"],
    "icon_properties": ["properties", "object properties", "file properties"],
    "icon_quick_part": ["quick part", "insert quick part", "building block"],
    "icon_read_aloud": ["read aloud", "text to speech", "listen document"],
    "icon_read_mode": ["read mode", "focus mode", "reading view"],
    "icon_replace": ["replace", "find and replace", "replace text"],
    "icon_restrict_editing": ["restrict editing", "limit editing", "protect document"],
    "icon_review_comments": ["review comments", "check comments", "comment review"],
    "icon_reviewing_pane": ["reviewing pane", "revision pane", "changes pane"],
    "icon_rotate": ["rotate", "rotate object", "rotate image"],
    "icon_ruler": ["ruler", "show ruler", "display ruler"],
    "icon_screenshot": ["screenshot", "take screenshot", "capture screen"],
    "icon_select": ["select", "highlight text", "choose object"],
    "icon_selection_pane": ["selection pane", "object selection", "layer pane"],
    "icon_set_default": ["set default", "default style", "make default"],
    "icon_shading": ["shading", "cell shading", "background color"],
    "icon_shapes": ["shapes", "insert shapes", "drawing shapes"],
    "icon_show_comments": ["show comments", "display comments", "review comments"],
    "icon_show_hide_p": ["show/hide paragraph marks", "display non-printing characters"],
    "icon_show_markup": ["show markup", "track changes markup", "revision markup"],
    "icon_show_traning": ["training", "tutorial", "help guide"],
    "icon_signature_line": ["signature line", "add signature", "insert signature"],
    "icon_smart_art": ["smart art", "insert smart art", "diagram"],
    "icon_spelling_grammer": ["spell check", "grammar check", "proofing"],
    "icon_split": ["split", "split document", "divide document"],
    "icon_strikethrough": ["strikethrough", "strike text", "line through text"],
    "icon_subscript": ["subscript", "lower text", "subscript text"],
    "icon_superscript": ["superscript", "raise text", "superscript text"],
    "icon_switch_window": ["switch window", "toggle windows", "change window"],
    "icon_symbol": ["symbol", "insert symbol", "special character"],
    "icon_table": ["table", "insert table", "add table","create table"],
    "icon_table_of_contents": ["table of contents", "toc", "insert toc"],
    "icon_text_box": ["text box", "insert text box", "add textbox"],
    "icon_text_color": ["text color", "font color"],
    "icon_text_effect": ["text effect", "apply effect", "effect text"],
    "icon_themes": ["themes", "apply theme", "document theme"],
    "icon_thesaurus": ["thesaurus", "synonyms", "find synonym"],
    "icon_track_changes": ["track changes", "revision tracking", "enable track changes"],
    "icon_translate": ["translate", "translate text", "language translation"],
    "icon_underline": ["underline", "underline text", "underscore"],
    "icon_update_table": ["update table", "refresh table", "table update"],
    "icon_view_ruler": ["view ruler", "show ruler", "ruler display"],
    "icon_watermark": ["watermark", "add watermark", "document watermark"],
    "icon_web_layout": ["web layout", "web view", "online layout"],
    "icon_word_art": ["word art", "insert word art", "stylized text"],
    "icon_word_count": ["word count", "count words", "document statistics"],
    "icon_wrap_text": ["wrap text", "text wrapping", "wrap around image"],
    "icon_zoom": ["zoom", "zoom in", "zoom out", "magnify"],
    "icon_zoom_100": ["zoom 100%", "reset zoom", "default zoom"],
    "lasso_select": ["lasso select", "draw select", "select area"],
    "layout_align": ["layout align", "align objects", "align elements"],
    "layout_indent": ["layout indent", "increase indent", "decrease indent"],
    "layout_spacing": ["layout spacing", "space between elements", "adjust spacing"],
    "styles": ["styles", "apply style", "format style",'text style'],
    "tab_design_active": ["design tab", "active design", "design active"],
    "tab_design_inactive": ["design tab inactive", "inactive design"],
    "tab_draw_active": ["draw tab active", "drawing active","tab drawing"],
    "tab_draw_inactive": ["draw tab inactive", "drawing inactive"],
    "tab_file_inactive": ["file tab inactive", "file menu inactive"],
    "tab_help_active": ["help tab active", "help menu active","tab help"],
    "tab_help_inactive": ["help tab inactive"],
    "tab_home_active": ["home tab active", "home menu active", "home tab","tab home"],
    "tab_home_inactive": ["home tab inactive"],
    "tab_insert_active": ["insert tab active", "insert menu active","tab insert"],
    "tab_insert_inactive": ["insert tab inactive"],
    "tab_layout_active": ["layout tab active", "layout menu active","tab layout"],
    "tab_layout_inactive": ["layout tab inactive"],
    "tab_mailings_inactive": ["mailings tab inactive"],
    "tab_references_active": ["references tab active", "references menu","tab reference"],
    "tab_references_inactive": ["references tab inactive"],
    "tab_review_active": ["review tab active", "review menu","tab review"],
    "tab_review_inactive": ["review tab inactive"],
    "tab_view_active": ["view tab active", "view menu","tab view"],
    "tab_view_inactive": ["view tab inactive"],
        }


    # Office-specific system prompt for Gemma
        self.system_prompt = (
            "You are a Microsoft Word and Excel built-in assistant. "
            "When given a user request, identify **which Office ribbon tab** "
            "(e.g., Home, Insert, Design, Layout, References, Review, View, Draw, etc.) "
            "contains the feature or option needed to complete that task.\n\n"
            "Respond strictly in JSON format like:\n"
            "{ \"tab\": \"Insert\" }\n\n"
            "Only respond with one most relevant tab name. "
            "Do NOT include any explanation, extra text, or notes. "
            "No markdown, no thinking text."

        )

    def query(self, user_text: str) -> dict:
        """Main query to identify which tab the user's request belongs to."""
        label = self._match_label(user_text)
        tab_data = self._generate_tab_with_gemma(user_text)
        return {"intent": user_text, "label": label, "tabs": tab_data}

    def _match_label(self, user_text: str):
        user_text = user_text.lower()
        best_label = None
        best_score = 0
        for label in self.labels:
            keywords = self.semantic_hints.get(label, [])
            for kw in keywords:
                score = difflib.SequenceMatcher(None, user_text, kw.lower()).ratio()
                if score > best_score:
                    best_score = score
                    best_label = label
        return best_label
    
    def _generate_tab_with_gemma(self, user_text: str):
        """
        Use Gemma to find the correct Microsoft Word tab (e.g., Home, Insert, Layout).
        Returns only the tab name as a plain string.
        """

        prompt = (
            f"{self.system_prompt}\n\n"
            f"User: \"{user_text}\"\n"
            f"Return ONLY valid JSON in this exact format: {{\"tab\": \"Home\"}}"
        )

        try:
            proc = subprocess.run(
                ["ollama", "run", self.model_name],
                input=prompt.encode("utf-8"),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=25
            )

            output = proc.stdout.decode("utf-8", errors="ignore").strip()
            output = re.sub(r"(?i)thinking.*?(?=\n|$)", "", output)

            match = re.search(r"\{.*\}", output, re.DOTALL)
            if match:
                try:
                    parsed = json.loads(match.group(0))
                    return parsed.get("tab", "").strip()
                except json.JSONDecodeError:
                    pass

            # Fallback: extract tab keyword from text
            for tab in [
                "Home", "Insert", "Design", "Layout", "References",
                "Review", "View", "Draw", "Mailings", "Help"
            ]:
                if tab.lower() in output.lower():
                    return tab

            return ""

        except subprocess.TimeoutExpired:
            return ""
        except Exception:
            return ""