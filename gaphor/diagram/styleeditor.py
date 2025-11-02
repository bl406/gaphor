from gi.repository import Gdk, Gtk, Gio, GLib, Adw
from pathlib import Path
from gaphor.ui.utils import state
from gaphor.core import event_handler, gettext
from gaphor.core.modeling import (
    AttributeUpdated,
    ElementDeleted,
    Presentation,
    StyleSheet,
    css_name,
)
from gaphor.ui import utils
from gaphor.core.modeling.diagram import StyledItem
from gaphor.core.styling import Color, Style
from gaphor.diagram.propertypages import PropertyPageBase, PropertyPages, new_builder
from gaphor.i18n import translated_ui_string
from gaphor.transaction import Transaction
from importlib.resources import files


class ImagesViewerWindow(Gtk.Window):
    def __init__(self, parent_window: Gtk.Window, datalist, query: str):
        super().__init__(title=f"图片检索：{query} ('共{len(datalist)}张图片')")
        self.set_transient_for(parent_window)
        self.set_default_size(900, 600)

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

        # === 逐条渲染：图片 + 描述 ===
        for tup in datalist:
            if not isinstance(tup, (tuple, list)) or len(tup) < 3:
                continue
            img_bytes, caption, img_id = tup

            # 将 bytes 转换为 Gdk.Texture（无须 GdkPixbuf）
            texture = None
            try:
                texture = Gdk.Texture.new_from_bytes(GLib.Bytes.new(img_bytes))
            except Exception:
                # 跳过坏图
                continue

            # 显示图片：用 Gtk.Picture，自动缩放不失真
            picture = Gtk.Picture.new_for_paintable(texture)
            picture.set_halign(Gtk.Align.CENTER)
            picture.set_can_shrink(True)
            picture.set_content_fit(Gtk.ContentFit.SCALE_DOWN)
            picture.set_size_request(200, 300)
            box.append(picture)

            # 显示描述（在图片下面）
            text_parts = []
            if isinstance(caption, str) and caption.strip():
                text_parts.append(caption.strip())
            text_parts.append(f"(ID: {img_id})")
            label = Gtk.Label(label="  ".join(text_parts))
            label.set_wrap(True)
            label.set_xalign(0.0)
            label.set_halign(Gtk.Align.CENTER)
            box.append(label)


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
        # 新增：show-images 的 Builder
        self.propertypages_builder_image = new_builder(
            "show-images",
            signals={"open-show-images": (self._on_open_show_images,)}
        )
        

    def construct(self):
        if not self.subject:
            return
        assert self.watcher
        # 返回“样式编辑”那块已有的 widget（保持原行为）
        editor_box = self.propertypages_builder.get_object("style-editor")
        
        # 取到我们新加的“Show images”按钮区块
        images_box = self.propertypages_builder_image.get_object("show-images")
        
        # 把两个块垂直叠加（也可以改为只返回 images_box，看你想怎么排）
        vbox = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        vbox.append(images_box)
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

    def _on_open_show_images(self, _button):
        # 取当前块的“名称”（如果有 subject 就优先用 subject.name）
        title = getattr(getattr(self.subject, "subject", None), "name", None)
        if not title:
            # 退化：拿呈现项的类型名
            title = type(self.subject).__name__
        # 打开“图片展示”窗口
        tree = state.image_tree
        datalist = utils.img_show.get_datalist_by_text(tree, title)
        win = ImagesViewerWindow(
            parent_window=self.main_window.window,
            datalist=datalist,
            query=title
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
