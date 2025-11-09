from gi.repository import Gdk, Gtk, Gio, GLib, Adw
from pathlib import Path
from gaphor.core import event_handler, gettext
from gaphor.core.modeling import (
    AttributeUpdated,
    ElementDeleted,
    Presentation,
    StyleSheet,
    css_name,
)
from gaphor.core.modeling.diagram import StyledItem
from gaphor.core.styling import Color, Style
from gaphor.diagram.propertypages import PropertyPageBase, PropertyPages, new_builder
from gaphor.i18n import translated_ui_string
from gaphor.transaction import Transaction
from importlib.resources import files

from gaphor.ui import utils
from gaphor.ui.utils import state, Acsemai
from gaphor.ui.utils.table_tree import wcTable
from docx import Document

class ImagesViewerWindow(Gtk.Window):
    def __init__(self, parent_window: Gtk.Window, datalist, title: str):
        super().__init__(title=title)
        self.set_transient_for(parent_window)
        self.set_default_size(350, 680)
        self.datalist = datalist

        # 滚动容器（纵向滚动）
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.ALWAYS)

        # 内容容器（纵向排布）
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        box.set_margin_top(12)
        box.set_margin_bottom(12)
        box.set_margin_start(12)
        box.set_margin_end(12)
        scrolled.set_child(box)
        self.set_child(scrolled)

        self.box = box
        self._populate_list(self.datalist)
        
    def _populate_list(self, items):
        """把 datalist 渲染为纵向的按钮列表；每个按钮显示 ImageCaption，点击回调."""
        # self._clear_box(self._content_box)
        if not items:
            empty = Gtk.Label(label="未找到匹配的图片")
            empty.set_xalign(0.0); empty.set_wrap(True)
            self.box.append(empty)
            return

        for i, data in enumerate(items):
            # data: ('table', wcTable, '表1 最大额定值', 0)
            caption = str(data[2])

            btn = Gtk.Button(label=caption)
            # 让按钮更紧凑/扁平（可选）
            btn.add_css_class("pill")
            btn.set_halign(Gtk.Align.FILL)
            btn.set_hexpand(True)
            btn.set_margin_start(20)
            btn.set_margin_end(20)

            # 点击回调：把这条 data 原样传回去
            def on_clicked(_b, payload=data):
                self.on_item_activated(payload)

            btn.connect("clicked", on_clicked)
            self.box.append(btn)
    
    def on_item_activated(self, data):
        """
        data 形如: ('table', wcTable, '表1 最大额定值', table_index)
        """
        _kind, img, caption, img_index = data

        # 1) 取应用 / 父窗
        parent = self
        app = parent.get_application() if parent else None

        # 3) 打开编辑窗口（标题可带表名）
        title = f"图片查看：{caption}"
        win = ImageEditorWindow(
            parent_window=self,
            img_tup=data,
            title=title
            )
        win.present()

class TablesViewerWindow(Gtk.ApplicationWindow):
    """
    表格检索/展示窗（GTK4）
    - 与 Application/父窗体绑定
    - 头部：SearchEntry（预留钩子）
    - 中部：滚动内容区（你来填充）
    - 底部：状态栏（可显示计数/提示）
    - 快捷键：Esc / Ctrl+W 关闭；Ctrl+F 聚焦搜索框
    """
    def __init__(self, parent_window: Gtk.Window, datalist, title: str):
        app = parent_window.get_application() if parent_window else None

        super().__init__(application=app, title=title)
        if parent_window:
            self.set_transient_for(parent_window)

        self.set_default_size(350, 680)

        # ---- 数据 ----
        self.datalist = datalist or []

        # ---- 顶层布局：垂直 Box ----
        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        root.set_margin_top(8); root.set_margin_bottom(8)
        root.set_margin_start(10); root.set_margin_end(10)
        self.set_child(root)

        # ========== 中部：内容区（你来塞组件） ==========
        # 滚动容器（纵向滚动）
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.ALWAYS)

        # 内容容器（纵向排布）
        self.box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        self.box.set_margin_top(12)
        self.box.set_margin_bottom(12)
        self.box.set_margin_start(12)
        self.box.set_margin_end(12)
        scrolled.set_child(self.box)
        self.set_child(scrolled)

        # 显示表格按钮
        self._populate_list(self.datalist)
    
    def _clear_box(self, box: Gtk.Box):
        """清空 box 子控件（GTK4 遍历）"""
        child = box.get_first_child()
        while child is not None:
            nxt = child.get_next_sibling()
            box.remove(child)
            child = nxt

    def _populate_list(self, items):
        """把 datalist 渲染为纵向的按钮列表；每个按钮显示 TableCaption，点击回调."""
        # self._clear_box(self._content_box)
        if not items:
            empty = Gtk.Label(label="未找到匹配的表格")
            empty.set_xalign(0.0); empty.set_wrap(True)
            self.box.append(empty)
            return

        for i, data in enumerate(items):
            # data: ('table', wcTable, '表1 最大额定值', 0)
            caption = str(data[2])

            btn = Gtk.Button(label=caption)
            # 让按钮更紧凑/扁平（可选）
            btn.add_css_class("pill")
            btn.set_halign(Gtk.Align.FILL)
            btn.set_hexpand(True)
            btn.set_margin_start(20)
            btn.set_margin_end(20)

            # 点击回调：把这条 data 原样传回去
            def on_clicked(_b, payload=data):
                self.on_item_activated(payload)

            btn.connect("clicked", on_clicked)
            self.box.append(btn)
    
    def on_item_activated(self, data):
        """
        data 形如: ('table', wcTable, '表1 最大额定值', table_index)
        """
        _kind, wct, caption, tbl_index = data

        # 1) 取应用 / 父窗
        parent = self
        app = parent.get_application() if parent else None
        docx_path = state.docx_path

        # 3) 打开编辑窗口（标题可带表名）
        title = f"表格编辑：{caption}"
        win = TableEditorWindow(
            parent_window=self,
            wc_table=wct,
            table_index=int(tbl_index),
            docx_path=docx_path,
            title=title,
        )
        win.present()
        
# --- 小工具：GTK4 清空容器 ---
def _clear_container(container: Gtk.Widget):
    child = container.get_first_child()
    while child is not None:
        nxt = child.get_next_sibling()
        container.remove(child)
        child = nxt

# --- 小工具：视觉列号（展开横向合并）---
def _visual_col_of(wctable: "wcTable", r: int, cidx: int) -> int:
    col = 0
    for j in range(cidx):
        span = wctable[r][j].grid_span or 1
        col += span
    return col


class TableEditorWindow(Gtk.ApplicationWindow):
    """
    可编辑的表格窗口（GTK4）：
      - 显示：将 wcTable 渲染到 Gtk.Grid（仅横向合并）
      - 保存：把修改写回当前 Document 的原 Cell，再保存到 docx_path
      - 重建：从 docx_path 重载后重建 UI（用于外部变更后刷新）
    依赖：传入 docx_path（保存所需）
    """
    def __init__(self, parent_window: Gtk.Window, wc_table: "wcTable",
                 table_index: int, docx_path: str, title: str = None):
        app = parent_window.get_application() if parent_window else None
        super().__init__(application=app, title=title or "表格编辑")
        if parent_window:
            self.set_transient_for(parent_window)
        self.set_default_size(1100, 720)

        # ---- 状态 ----
        self.docx_path = docx_path
        self.table_index = table_index
        self.doc = None               # 当前 Document
        self.wct: "wcTable" = wc_table    # 当前 wcTable（与 self.doc 对应）
        self.grid: Gtk.Grid = None
        self.patches = {}             # {cell_id: new_text}
        self.id_to_cell = {}          # {cell_id: python-docx _Cell}

        # ---- UI 框架 ----
        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        root.set_margin_top(8); root.set_margin_bottom(8)
        root.set_margin_start(10); root.set_margin_end(10)
        self.set_child(root)

        # 顶部工具条（保存 / 重建 / AI）
        btn_save = Gtk.Button(label="保存回 DOCX")
        btn_save.connect("clicked", self._on_save_clicked)

        btn_rebuild = Gtk.Button(label="重建（从文件重载）")
        btn_rebuild.connect("clicked", self._on_rebuild_clicked)
        
        btn_Acsemai = Gtk.Button(label="打开 AcsemAI 智能问答模型")
        btn_Acsemai.add_css_class("suggested-action")
        btn_Acsemai.connect("clicked", self._on_Acsemai_clicked)

        topbar = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        topbar.append(btn_save)
        topbar.append(btn_rebuild)
        topbar.append(btn_Acsemai)
        root.append(topbar)

        # 表格滚动区
        self.scrolled = Gtk.ScrolledWindow()
        self.scrolled.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        root.append(self.scrolled)

        # 初始化：装载文档 & 渲染
        self._load_and_build(wc_table)

    # ---------- 装载 & 渲染 ----------
    def _load_and_build(self, wc_table_from_list: "wcTable" = None):
        """
        第一次：如果传入 wc_table_from_list（来自列表搜索），先用它渲染，
        但为了保存需要 docx_path，对齐起见这里直接从 docx_path 重载 Document + wcTable。
        """
        if not self.docx_path:
            # 兜底：如果没传 docx_path，禁用保存，并用传入的 wct 渲染（只读）
            self.doc = None
            self.wct = wc_table_from_list
        else:
            self.doc = Document(self.docx_path)
            self.wct = wcTable(self.doc.tables[self.table_index], self.table_index)

        self._build_grid_from_wct()

    def _build_grid_from_wct(self):
        if self.grid is None:
            self.grid = Gtk.Grid(column_spacing=6, row_spacing=6,
                                 margin_top=12, margin_bottom=12,
                                 margin_start=12, margin_end=12)
            self.scrolled.set_child(self.grid)
        else:
            _clear_container(self.grid)

        self.patches.clear()
        self.id_to_cell.clear()

        # 逐行渲染
        for r in range(self.wct._rows):
            row_cells = self.wct[r]
            if not row_cells:
                continue

            for cidx, wcc in enumerate(row_cells):
                vcol = _visual_col_of(self.wct, r, cidx)
                colspan = wcc.grid_span or 1

                buf = Gtk.TextBuffer()
                buf.set_text(wcc.text or "")

                tv = Gtk.TextView(buffer=buf)
                tv.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
                tv.set_vexpand(True)
                tv.set_hexpand(True)
                tv.set_size_request(160, 64)

                # 唯一 ID，沿用你 wcTable 的 title（t{tbl}_r{r}_c{cidx}）
                cell_id = wcc.title
                tv.set_name(cell_id)

                # 保存映射：写回用
                self.id_to_cell[cell_id] = wcc.cell

                def on_changed(buffer, cid=cell_id):
                    start, end = buffer.get_bounds()
                    self.patches[cid] = buffer.get_text(start, end, True)
                buf.connect("changed", on_changed)

                # 仅横向合并：rowspan = 1
                self.grid.attach(tv, vcol, r, colspan, 1)

    # ---------- 保存/重建 ----------
    def _on_save_clicked(self, _btn):
        if not self.docx_path or not self.doc:
            print("⚠ 当前窗口缺少 docx_path 或 Document，无法保存。")
            return
        if not self.patches:
            print("无修改")
            return

        # 根据 patches 写回 python-docx 的原始 Cell
        for cid, txt in self.patches.items():
            cell = self.id_to_cell.get(cid)
            if not cell:
                continue
            # 清空并按行重写
            cell.text = ""
            lines = (txt or "").split("\n")
            if lines:
                cell.paragraphs[0].text = lines[0]
                for extra in lines[1:]:
                    p = cell.add_paragraph()
                    p.text = extra

        self.doc.save(self.docx_path)
        print(f"✅ 已保存：{self.docx_path}")
        self.patches.clear()

    def _on_rebuild_clicked(self, _btn):
        if not self.docx_path:
            print("⚠ 没有 docx_path，无法从文件重载。")
            return
        self.doc = Document(self.docx_path)
        self.wct = wcTable(self.doc.tables[self.table_index], self.table_index)
        self._build_grid_from_wct()
        print("🔁 已从文件重载并重建 UI")
        
    def _on_Acsemai_clicked(self, _btn):
        title = "AcsemAI 智能问答模型"
        win = AcsemAITableChatWindow(
            parent_window=self,
            title=title,
            table = self.wct
        )
        win.present()
   
class ImageEditorWindow(Gtk.ApplicationWindow):
    def __init__(self, parent_window: Gtk.Window, img_tup, title: str = None):
        app = parent_window.get_application() if parent_window else None
        super().__init__(application=app, title=title or "图片查看")
        if parent_window:
            self.set_transient_for(parent_window)
        self.set_default_size(1100, 720)

        # ---- 状态 ----
        # 期望 img_tup: (tag, img_bytes, caption, img_id)
        self.tag, self.img_bytes, self.img_caption, self.img_id = img_tup
        self._texture = Gdk.Texture.new_from_bytes(GLib.Bytes.new(self.img_bytes))
        self._orig_w = self._texture.get_width()
        self._orig_h = self._texture.get_height()
        self._scale = 1.0
        self._min_scale = 0.1
        self._max_scale = 8.0

        # ---- 顶层布局 ----
        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        root.set_margin_top(8); root.set_margin_bottom(8)
        root.set_margin_start(10); root.set_margin_end(10)
        self.set_child(root)

        # 顶部工具条
        topbar = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        root.append(topbar)

        btn_ai = Gtk.Button(label="打开 AcsemAI 智能问答模型")
        btn_ai.add_css_class("suggested-action")
        btn_ai.connect("clicked", self._on_acsemai_clicked)
        topbar.append(btn_ai)

        topbar.append(Gtk.Separator(orientation=Gtk.Orientation.VERTICAL))

        btn_zoom_out = Gtk.Button(label="−")
        btn_zoom_out.set_tooltip_text("缩小 (Ctrl+滚轮下 / -)")
        btn_zoom_out.connect("clicked", lambda _b: self._zoom_step(-0.1))
        topbar.append(btn_zoom_out)

        btn_zoom_in = Gtk.Button(label="+")
        btn_zoom_in.set_tooltip_text("放大 (Ctrl+滚轮上 / +)")
        btn_zoom_in.connect("clicked", lambda _b: self._zoom_step(+0.1))
        topbar.append(btn_zoom_in)

        btn_zoom_100 = Gtk.Button(label="100%")
        btn_zoom_100.set_tooltip_text("实际尺寸 (0)")
        btn_zoom_100.connect("clicked", lambda _b: self._set_scale(1.0))
        topbar.append(btn_zoom_100)

        btn_zoom_fit = Gtk.Button(label="适配")
        btn_zoom_fit.set_tooltip_text("适配窗口")
        btn_zoom_fit.connect("clicked", self._on_fit_clicked)
        topbar.append(btn_zoom_fit)

        # ---- 中部可滚动区 ----
        self.scrolled = Gtk.ScrolledWindow()
        self.scrolled.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        self.scrolled.set_size_request(800, 600)

        # 内容：图片 + 说明标签（纵向）
        self.content_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        self.content_box.set_margin_top(12)
        self.content_box.set_margin_bottom(12)
        self.content_box.set_margin_start(12)
        self.content_box.set_margin_end(12)
        self.scrolled.set_child(self.content_box)
        root.append(self.scrolled)

        # 图片
        self.picture = Gtk.Picture.new_for_paintable(self._texture)
        self.picture.set_halign(Gtk.Align.CENTER)
        self.picture.set_can_shrink(True)
        self.picture.set_content_fit(Gtk.ContentFit.SCALE_DOWN)  # 初始：缩小以适配
        # 先给一个合理的最小可见区域
        self.picture.set_size_request(min(self._orig_w, 600), min(self._orig_h, 400))
        self.content_box.append(self.picture)

        # 说明文字
        parts = []
        if isinstance(self.img_caption, str) and self.img_caption.strip():
            parts.append(self.img_caption.strip())
        parts.append(f"(ID: {self.img_id})")
        self.caption_label = Gtk.Label(label="  ".join(parts))
        self.caption_label.set_wrap(True)
        self.caption_label.set_xalign(0.5)
        self.caption_label.set_halign(Gtk.Align.CENTER)
        self.content_box.append(self.caption_label)

        # ---- 键鼠缩放支持 ----
        # Ctrl + 滚轮缩放
        self._ctrl_scroller = Gtk.EventControllerScroll.new(Gtk.EventControllerScrollFlags.VERTICAL)
        self._ctrl_scroller.connect("scroll", self._on_scroll_zoom)
        self.add_controller(self._ctrl_scroller)

        # 键盘 +/-/0
        key = Gtk.EventControllerKey()
        key.connect("key-pressed", self._on_key_pressed)
        self.add_controller(key)

        # 初始尝试适配（等窗口布局稳定后再算）
        GLib.idle_add(lambda: (self._fit_to_view(), False))

    # ====== 行为 ======

    def _on_acsemai_clicked(self, _btn):
        title = "AcsemAI 智能问答模型"
        win = AcsemAIImageChatWindow(
            parent_window=self,
            title=title,
            image_blob=self.img_bytes,
            image_caption=self.img_caption
        )
        win.present()

    def _zoom_step(self, delta):
        self._set_scale(self._scale * (1.0 + delta))

    def _set_scale(self, s: float):
        s = max(self._min_scale, min(self._max_scale, s))
        self._scale = s
        # 缩放通过调整 size_request 实现滚动查看
        new_w = max(1, int(self._orig_w * self._scale))
        new_h = max(1, int(self._orig_h * self._scale))
        self.picture.set_content_fit(Gtk.ContentFit.FILL)  # 缩放时用尺寸驱动
        self.picture.set_size_request(new_w, new_h)

    def _fit_to_view(self):
        # 根据当前可视区域估算一个“适配窗口”的缩放
        alloc = self.scrolled.get_allocation()
        view_w = max(1, alloc.width - 64)   # 留点边距，避免贴边
        view_h = max(1, alloc.height - 120) # 顶部工具条 + 说明标签的空间
        if view_w <= 0 or view_h <= 0:
            return
        scale_w = view_w / self._orig_w
        scale_h = view_h / self._orig_h
        self._set_scale(min(scale_w, scale_h))

    def _on_fit_clicked(self, _btn):
        self._fit_to_view()

    def _on_scroll_zoom(self, controller, dx, dy):
        # 仅当按住 Ctrl 时滚轮缩放；否则让滚动窗处理
        # 说明：GTK4 在 key-pressed 里记录不到 Ctrl 状态给 scroll，所以简单做法：
        # 用 Shift 反向也可：向上放大，向下缩小
        if dy is None:
            return False
        # 这里无修饰键就直接做“自然缩放”：上滚放大，下滚缩小
        if dy < 0:
            self._zoom_step(+0.1)
            return True
        elif dy > 0:
            self._zoom_step(-0.1)
            return True
        return False

    def _on_key_pressed(self, _ctrl, keyval, _keycode, _state):
        if keyval in (Gdk.KEY_plus, Gdk.KEY_KP_Add):
            self._zoom_step(+0.1); return True
        if keyval in (Gdk.KEY_minus, Gdk.KEY_KP_Subtract):
            self._zoom_step(-0.1); return True
        if keyval in (Gdk.KEY_0, Gdk.KEY_KP_0):
            self._set_scale(1.0); return True
        return False
        
class AcsemAITableChatWindow(Gtk.ApplicationWindow):
    def __init__(self, parent_window: Gtk.Window, title, table):
        app = parent_window.get_application()
        super().__init__(application=app, title=title)
        if parent_window:
            self.set_transient_for(parent_window)
        self.set_default_size(840, 550)
        self.parent_window = parent_window
        self.wct = table

        # ---- 顶层布局：垂直 Box ----
        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        root.set_margin_top(8); root.set_margin_bottom(8)
        root.set_margin_start(10); root.set_margin_end(10)
        self.set_child(root)

        # ========== 顶部：输入+按钮（整体居中） ==========
        controls = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        controls.set_halign(Gtk.Align.CENTER)   # ⭐ 整组居中
        controls.set_valign(Gtk.Align.START)

        # 输入框放在一个专用的滚动容器里（长文本时滚动）
        self.input_view = Gtk.TextView()
        self.input_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)   # 自动换行
        self.input_view.set_top_margin(6)
        self.input_view.set_bottom_margin(6)
        self.input_view.set_left_margin(8)
        self.input_view.set_right_margin(8)
        self.input_view.add_css_class("search-entry") 

        self.input_scrolled = Gtk.ScrolledWindow()
        self.input_scrolled.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)  # 只纵向滚
        self.input_scrolled.add_css_class("search-entry")  
        self.input_scrolled.set_child(self.input_view)
        # 设定一个“视觉上合适”的宽高（固定高度，超出就滚动）
        self.input_scrolled.set_size_request(840, 100)
        
        # 输出框放在一个专用的滚动容器里（长文本时滚动）
        self.output_view = Gtk.TextView()
        self.output_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)   # 自动换行
        self.output_view.set_top_margin(6)
        self.output_view.set_bottom_margin(6)
        self.output_view.set_left_margin(8)
        self.output_view.set_right_margin(8)
        self.output_view.add_css_class("search-entry") 

        self.output_scrolled = Gtk.ScrolledWindow()
        self.output_scrolled.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)  # 只纵向滚
        self.output_scrolled.add_css_class("search-entry")  
        self.output_scrolled.set_child(self.output_view)
        # 设定一个“视觉上合适”的宽高（固定高度，超出就滚动）
        self.output_scrolled.set_size_request(840, 450)
        
        # 执行按钮
        run_btn = Gtk.Button(label="发送")
        run_btn.add_css_class("search-run")
        run_btn.connect("clicked", self._on_run_clicked)
        
        key = Gtk.EventControllerKey()
        key.connect("key-pressed", self._on_input_key_pressed)
        self.input_view.add_controller(key)

        controls.append(self.input_scrolled)
        controls.append(self.output_scrolled)
        controls.append(run_btn)
        root.append(controls)
        
    # 读取输入文本
    def _get_input_text(self) -> str:
        buf = self.input_view.get_buffer()
        start, end = buf.get_bounds()
        return buf.get_text(start, end, True)

    # 写入结果文本
    def set_output_text(self, text: str):
        buf = self.output_view.get_buffer()
        buf.set_text(text)

    # 点击“执行”后的行为
    def _on_run_clicked(self, _button):
        text = self._get_input_text().strip()
        print("[AcsemAIWindow] run with:", text)
        result = Acsemai.table_chat(text, self.wct)
        print("[AcsemAIWindow] answer as:", result)
        self.set_output_text(result)
        return 
    
    # 回车执行
    def _on_input_key_pressed(self, controller, keyval, keycode, state):
        if keyval in (Gdk.KEY_Return, Gdk.KEY_KP_Enter):
            self._on_run_clicked(None)
            # ✅ 返回 True 表示我们处理了这个按键，
            # TextView 不应该继续接收这一回车（避免换行）
            return True
        return False
    
class AcsemAIImageChatWindow(AcsemAITableChatWindow):
    def __init__(self, parent_window: Gtk.Window, title, image_blob, image_caption):
        super().__init__(parent_window=parent_window, title=title, table=None)
        self.image_tup = (image_blob, image_caption)
    # 点击“执行”后的行为
    def _on_run_clicked(self, _button):
        text = self._get_input_text().strip()
        print("[AcsemAIWindow] run with:", text)
        result = Acsemai.image_chat(text, self.image_tup)
        print("[AcsemAIWindow] answer as:", result)
        self.set_output_text(result)
        return 
        
@PropertyPages.register(Presentation)
class StylePropertyPage(PropertyPageBase):
    """A button to open a easy-to-use CSS editor."""

    order = 450
    style_editor = None

    def __init__(self, subject, event_manager, element_factory, main_window):
        super().__init__()
        self.subject = subject
        self.event_manager = event_manager
        self.element_factory = element_factory
        self.main_window = main_window
        self.watcher = subject.watcher() if subject else None
        self.propertypages_builder = new_builder(
            "style-editor",
            signals={
                "open-style-editor": (self._on_open_style_editor,),
            },
        )
        
        # 将搜索图片和搜索表格合并
        self.propertypages_builder_tools = new_builder(
            "show-tools",
            signals={
                "open-show-images": (self._on_open_show_images,),
                "open-show-tables": (self._on_open_show_tables,),
            },
        )

    def construct(self):
        if not self.subject:
            return
        assert self.watcher
        
        #三个块：style-editor、show-images、show-tables
        editor_box = self.propertypages_builder.get_object("style-editor")
        tools_box = self.propertypages_builder_tools.get_object("show-tools")
        
        # 把横向容器和editor-box放入纵向容器
        vbox = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        vbox.append(tools_box)
        vbox.append(editor_box)
        return vbox
        # return self.propertypages_builder.get_object("style-editor")

    def _on_open_style_editor(self, button):
        if StylePropertyPage.style_editor:
            StylePropertyPage.style_editor.close()

        style_sheet = next(self.element_factory.select(StyleSheet))
        StylePropertyPage.style_editor = StyleEditor(
            self.subject,
            style_sheet,
            self.event_manager,
            self.main_window.window,
            self.close_style_editor,
        )
        StylePropertyPage.style_editor.present()

    # open-show-images回调函数
    def _on_open_show_images(self, _button):
        def _make_title(block_title, datalist):
            return f"图片检索：{block_title} ('共{len(datalist)}项')"
            
        # 取当前块的“名称”（如果有 subject 就优先用 subject.name）
        block_title = getattr(getattr(self.subject, "subject", None), "name", None)
        # 打开“图片展示”窗口
        tree = state.image_tree
        datalist = utils.img_show.get_datalist_by_text(tree, block_title)
        win = ImagesViewerWindow(
            parent_window=self.main_window.window,
            datalist=datalist,
            title=_make_title(block_title, datalist)
        )
        win.present()
    
    # open-show-tables回调函数
    def _on_open_show_tables(self, _button):
        def _make_title(block_title, datalist):
            return f"表格检索：{block_title} ('共{len(datalist)}项')"
               
        block_title = getattr(getattr(self.subject, "subject", None), "name", None)
        # 打开“表格展示”窗口
        tree = state.table_tree
        datalist = utils.table_tree.get_datalist_by_text(tree, block_title)
        
        win = TablesViewerWindow(
            parent_window=self.main_window.window,
            datalist=datalist,
            title=_make_title(block_title, datalist)
        )
        win.present()
    
    def close_style_editor(self):
        StylePropertyPage.style_editor = None


class StyleEditor:
    def __init__(
        self,
        subject: Presentation,
        style_sheet: StyleSheet,
        event_manager,
        parent_window,
        close_callback,
    ):
        self.subject = subject
        self.style_sheet = style_sheet
        self.event_manager = event_manager
        self.parent_window = parent_window
        self.close_callback = close_callback

        self.window_builder = Gtk.Builder()
        self.window_builder.add_from_string(
            translated_ui_string("gaphor.diagram", "styleeditor.ui")
        )

        self.window: Gtk.Window | None = None
        self.presentation_style: Style = {}

        event_manager.subscribe(self._on_name_changed)
        event_manager.subscribe(self._on_element_deleted)

    def present(self):
        if self.window is None:
            self.window = self.window_builder.get_object("style-editor")
            self.window.connect("close-request", self.close)
            self.color = self.window_builder.get_object("color")
            self.border_radius = self.window_builder.get_object("border-radius")
            self.background_color = self.window_builder.get_object("background-color")
            self.text_color = self.window_builder.get_object("text-color")
            self.window.set_transient_for(self.parent_window)

            self.fields()

            self.color.connect("notify::rgba", self.on_color_set)
            self.border_radius.connect("value-changed", self.on_border_radius_set)
            self.background_color.connect("notify::rgba", self.on_background_color_set)
            self.text_color.connect("notify::rgba", self.on_text_color_set)
            self.window_builder.get_object("export").connect("clicked", self.on_export)

        self.window.present()

    def fields(self):
        style = self.subject.diagram.style(StyledItem(self.subject))
        if style.get("color"):
            self.color.set_rgba(to_gdk_rgba(style["color"]))
            self.text_color.set_rgba(to_gdk_rgba(style["color"]))

        if style.get("border-radius"):
            self.border_radius.set_value(int(style["border-radius"]))

        if style.get("background-color"):
            self.background_color.set_rgba(to_gdk_rgba(style["background-color"]))

        if style.get("text-color"):
            self.text_color.set_rgba(to_gdk_rgba(style["text-color"]))

    def close(self, _widget=None):
        if self.window:
            self.window.destroy()
            self.window = None
        self.event_manager.subscribe(self._on_name_changed)
        self.style_sheet.instant_style_declarations = ""
        self.close_callback()

    def change_style(self, prop, value):
        self.presentation_style[prop] = value  # type: ignore[literal-required]
        self.style_sheet.instant_style_declarations = self.render_css()

    def on_color_set(self, widget, _paramspec=None):
        color = widget.get_rgba()
        self.change_style("color", from_gdk_rgba(color))

    def on_border_radius_set(self, widget):
        self.change_style("border-radius", widget.get_value())

    def on_background_color_set(self, widget, _paramspec=None):
        color = widget.get_rgba()
        self.change_style("background-color", from_gdk_rgba(color))

    def on_text_color_set(self, widget, _paramspec=None):
        color = widget.get_rgba()
        self.change_style("text-color", from_gdk_rgba(color))

    def on_export(self, _widget):
        with Transaction(self.event_manager):
            self.style_sheet.styleSheet += f"\n{self.render_css()}\n"
        self.close()

    @event_handler(AttributeUpdated)
    def _on_name_changed(self, event: AttributeUpdated):
        if event.property.name == "name":
            self.style_sheet.instant_style_declarations = self.render_css()

    @event_handler(ElementDeleted)
    def _on_element_deleted(self, event: ElementDeleted):
        if event.element is self.subject:
            self.close()

    def render_css(self) -> str:
        name = getattr(self.subject.subject, "name", None)
        tag = css_name(self.subject)
        selector = f'{tag}[name="{name}"]' if name else tag
        properties = "\n".join(
            f"  {k}: {v};" for k, v in self.presentation_style.items()
        )
        return f"{selector} {{\n{properties}\n}}"


def to_gdk_rgba(color: str | Color) -> Gdk.RGBA:
    rgba = Gdk.RGBA()
    if isinstance(color, str):
        rgba.parse(color)
    else:
        rgba.red, rgba.green, rgba.blue, rgba.alpha = color
    return rgba


def from_gdk_rgba(rgba: Gdk.RGBA) -> str:
    r = int(rgba.red * 255)
    g = int(rgba.green * 255)
    b = int(rgba.blue * 255)
    a = rgba.alpha
    return f"rgba({r}, {g}, {b}, {a})"
