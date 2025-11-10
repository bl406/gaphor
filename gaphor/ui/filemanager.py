"""The file service is responsible for loading and saving the user data."""

from __future__ import annotations

import logging
import shutil
import tempfile
import os
import asyncio
from gaphor.ui import utils
from gaphor.ui.utils import state, Acsemai
from gaphor.diagram.styleeditor import AcsemAIAllChatWindow, ImagesViewerWindow, TablesViewerWindow

from importlib.resources import files
from collections.abc import Callable
from functools import partial
from pathlib import Path

from gi.repository import Adw, Gio, Gtk, Gdk

import gaphor.storage as storage
from gaphor.abc import ActionProvider, Service
from gaphor.asyncio import sleep
from gaphor.babel import translate_model
from gaphor.core import action, event_handler, gettext
from gaphor.core.changeset.compare import compare
from gaphor.core.modeling import ElementFactory, ModelReady
from gaphor.event import (
    ModelSaved,
    Notification,
    SessionCreated,
    SessionShutdown,
    SessionShutdownRequested,
)
from gaphor.storage.mergeconflict import split_ours_and_theirs
from gaphor.storage.parser import MergeConflictDetected
from gaphor.ui.errordialog import error_dialog
from gaphor.ui.filedialog import GAPHOR_FILTER, save_file_dialog, open_file_dialog
from gaphor.ui.statuswindow import StatusWindow

import tarfile

READ_TEMPDIR = './read_temp'
WRITE_TEMPDIR = './write_temp'
DEFAULT_EXT = ".acsem"
DOC_EXT = ".acbin"
MAX_RECENT = 10

log = logging.getLogger(__name__)


def error_message(e):
    if not isinstance(e, IOError):
        return gettext(
            "ACSEM was not able to store the model, probably due to an internal error:\n{exc}\nIf you think this is a bug, please contact the developers."
        ).format(exc=str(e))
    if e.errno == 13:
        return gettext(
            "You do not have the permissions necessary to save the model.\nPlease check that you typed the location correctly and try again."
        )
    elif e.errno == 28:
        return gettext(
            "You do not have enough free space on the device to save the model.\nPlease free up some disk space and try again or save it in a different location."
        )
    return gettext(
        "The model cannot be stored at this location:\n{exc}\nPlease check that you typed the location correctly and try again."
    ).format(exc=str(e))

def delete_all_files_and_dirs(target_dir: Path):
    # 确保目标是目录
    if not target_dir.is_dir():
        raise ValueError(f"{target_dir} 不是有效目录")
    
    # 递归处理子目录：先删除子目录内的所有内容
    for item in target_dir.iterdir():
        if item.is_file():
            # 删除文件
            item.unlink()            
        elif item.is_dir():
            # 递归删除子目录（先删子目录内的内容，再删自身）
            delete_all_files_and_dirs(item)
            # 子目录内容删除后，删除空目录
            item.rmdir()

class FileManager(Service, ActionProvider):
    """The file service, responsible for loading and saving Gaphor models."""

    def __init__(self, event_manager, element_factory, modeling_language, main_window):
        """File manager constructor.

        There is no current filename yet.
        """
        super().__init__()
        self.event_manager = event_manager
        self.element_factory = element_factory
        self.modeling_language = modeling_language
        self.main_window = main_window
        self._filename: Path | None = None
        self._monitor: Gio.Monitor | None = None
        self._filename_realread: Path | None = None

        event_manager.subscribe(self._on_session_shutdown_request)
        event_manager.subscribe(self._on_session_created)
    
    ### 构建图片树
    async def _build_image_tree_background(self, filename: Path) -> None:
        try:
            docx_file = filename.with_suffix('.acbin')
            print(docx_file)
            tree = await asyncio.to_thread(utils.img_show.Create_image_tree, docx_file)
            state.image_tree = tree
        except Exception:
            log.exception("Failed to build image tree for %s", filename)
            
    async def _build_table_tree_background(self, filename: Path) -> None:
        try:
            docx_file = filename.with_suffix('.acbin')
            print(docx_file)
            tree = await asyncio.to_thread(utils.table_tree.Create_table_tree, docx_file)
            state.table_tree = tree
        except Exception:
            log.exception("Failed to build table tree for %s", filename)

    def shutdown(self):
        """Called when shutting down the file manager service."""
        self.event_manager.unsubscribe(self._on_session_shutdown_request)
        self.event_manager.unsubscribe(self._on_session_created)

    @property
    def filename(self) -> Path | None:
        """Return the current file name.

        This method is used by the filename property.
        """
        return self._filename

    @filename.setter
    def filename(self, filename: Path | str | None):
        """Sets the current file name.

        This method is used by the filename property.
        """

        if filename != self._filename:
            self._filename = Path(filename) if filename else None

    def load_template(self, template):
        translated_model = translate_model(template)
        storage.load(translated_model, self.element_factory, self.modeling_language)
        self.event_manager.handle(ModelReady(self))

    async def load(self, filename: Path):
        """Load the Gaphor model from the supplied file name.

        A status window displays the loading progress. The load
        generator updates the progress queue.  The loader is passed to a
        GIdleThread which executes the load generator. If loading is
        successful, the filename is set.
        """
        # First claim file name, so any other files will be opened in a different session
        self.filename = filename       

        if filename.suffix == DEFAULT_EXT:
            extraction_folder = Path(READ_TEMPDIR).absolute()
            if extraction_folder.exists() == True:
                delete_all_files_and_dirs(extraction_folder)  # 清空目录
            # untar the file with filename to a local dir
            print(filename)
            with tarfile.open(filename, 'r') as tar:
                # 将所有内容解压到指定路径
                tar.extractall(path=extraction_folder)                

        status_window = StatusWindow(
            gettext("Loading…"),
            gettext("Loading model from {filename}").format(filename=filename),
            parent=self.parent_window,
        )

        try:
            filename_realread = filename
            print(filename)
            if filename.suffix == DEFAULT_EXT:
                file_stem = filename.stem  # 去掉扩展名后的文件名
                gaphor_fn = os.path.join(extraction_folder,file_stem,file_stem+'.gaphor')   
                filename_realread = Path(gaphor_fn)
                self._filename_realread = filename_realread
            print("filename really read:",filename_realread.name)
            await self._load_async(filename_realread, status_window.progress)
        finally:
            status_window.done()
        self.event_manager.handle(ModelReady(self, filename=filename_realread))
        state.docx_path = filename_realread.with_suffix('.acbin')
        asyncio.create_task(self._build_image_tree_background(filename_realread))
        asyncio.create_task(self._build_table_tree_background(filename_realread))


    @action("file-reload")
    async def reload(self):
        if self.filename and self.filename.exists():
            self.element_factory.flush()

            await self.load(self.filename)
            self.event_manager.handle(ModelReady(self))

    async def merge(
        self,
        ancestor_filename: Path,
        current_filename: Path,
        incoming_filename: Path,
    ):
        status_window = StatusWindow(
            gettext("Loading…"),
            gettext("Loading current and incoming model"),
            parent=self.parent_window,
        )

        def progress(percentage, completed=0):
            status_window.progress(completed + percentage / 3)

        try:
            log.debug("Loading current model from %s", current_filename)
            await self._load_async(current_filename, progress)

            log.debug("Loading ancestor model from %s", ancestor_filename)
            ancestor_element_factory = ElementFactory()
            await self._load_async(
                ancestor_filename,
                partial(progress, completed=33),
                element_factory=ancestor_element_factory,
            )

            log.debug("Loading incoming model from %s", incoming_filename)
            incoming_element_factory = ElementFactory()
            await self._load_async(
                incoming_filename,
                partial(progress, completed=66),
                element_factory=incoming_element_factory,
            )

            log.debug("Comparing models")
            with self.element_factory.block_events():
                list(
                    compare(
                        self.element_factory,
                        ancestor_element_factory,
                        incoming_element_factory,
                    )
                )
        finally:
            status_window.done()

    async def _load_async(
        self,
        filename: Path,
        progress: Callable[[int], None] | None = None,
        element_factory=None,
    ):
        factory = element_factory or self.element_factory
        try:
            with filename.open(encoding="utf-8", errors="replace") as file_obj:
                for percentage in storage.load_generator(
                    file_obj,
                    factory,
                    self.modeling_language,
                ):
                    if progress:
                        progress(percentage)
                    await sleep(0)
        except MergeConflictDetected:
            self.filename = None
            await self.resolve_merge_conflict(filename)
        except Exception as e:
            print("Exception during load:",e)
            self.filename = None
            await error_dialog(
                message=gettext("Unable to open model “{filename}”.").format(
                    filename=filename
                ),
                secondary_message=gettext(
                    "This file does not contain a valid Gaphor model."
                ),
                window=self.parent_window,
            )
            self.event_manager.handle(SessionShutdown())

    async def resolve_merge_conflict(self, filename: Path):
        temp_dir = tempfile.TemporaryDirectory()
        ancestor_filename = Path(temp_dir.name) / f"ancestor-{filename.name}"
        current_filename = Path(temp_dir.name) / f"current-{filename.name}"
        incoming_filename = Path(temp_dir.name) / f"incoming-{filename.name}"
        with (
            ancestor_filename.open("wb") as ancestor_file,
            current_filename.open("wb") as current_file,
            incoming_filename.open("wb") as incoming_file,
        ):
            split = split_ours_and_theirs(
                filename, ancestor_file, current_file, incoming_file
            )

        if split:
            answer = await resolve_merge_conflict_dialog(self.parent_window)
            if answer == "cancel":
                self.event_manager.handle(SessionShutdown())
            elif answer == "current":
                await self.load(current_filename)
            elif answer == "incoming":
                await self.load(incoming_filename)
            elif answer == "manual":
                await self.merge(ancestor_filename, current_filename, incoming_filename)
            else:
                raise ValueError(f"Unknown resolution for merge conflict: {answer}")

            temp_dir.cleanup()
            self.filename = filename
            self.event_manager.handle(
                ModelReady(self, filename=filename, modified=True)
            )
        else:
            await error_dialog(
                message=gettext("Unable to open model “{filename}”.").format(
                    filename=filename.name
                ),
                secondary_message=gettext(
                    "This file does not contain a valid Gaphor model."
                ),
                window=self.parent_window,
            )
            self.event_manager.handle(SessionShutdown())

    async def save(self, filename):
        """Save the current model to the specified file name.

        Before writing the model file, this will verify that there are
        no orphan references. It will also verify that the filename has
        the correct extension. A status window is displayed while the
        save operation is executed.
        """

        if not filename or (filename.exists() and not filename.is_file()):
            return

        status_window = (
            StatusWindow(
                gettext("Saving…"),
                gettext("Saving model to {filename}").format(filename=filename),
                parent=self.parent_window,
            )
            if self.element_factory.size() > 100
            else None
        )

        try:
            if filename.suffix != DEFAULT_EXT:
                with filename.open("w", encoding="utf-8") as out:
                    for percentage in storage.save_generator(out, self.element_factory):
                        if status_window:
                            status_window.progress(percentage)
                        await sleep(0)
                self.event_manager.handle(ModelSaved(filename))
            else:
                if not self._filename_realread.with_suffix(DOC_EXT):
                    await error_dialog(
                        message=gettext("无法保存模型"),
                        secondary_message=gettext("模型文件被损坏，无法找到保存所需的模型数据。"),
                        window=self.parent_window,
                    )
                    return # 目录下没找到我们的模型文件，无法保存  
                 
                export_folder = Path(WRITE_TEMPDIR).absolute()
                if export_folder.exists() == True:
                    delete_all_files_and_dirs(export_folder)  # 清空目录
                export_folder = export_folder.joinpath(self.filename.stem)
                export_folder.mkdir(parents=True, exist_ok=True)

                doc_filename = export_folder.joinpath(self._filename_realread.with_suffix(DOC_EXT).name)
                shutil.copy2(self._filename_realread.with_suffix(DOC_EXT), doc_filename)
                
                gaphor_file = os.path.join(export_folder,self.filename.stem+'.gaphor')                   
                await self.save(Path(gaphor_file))               
                
                # 打包成tar文件
                with tarfile.open(filename, 'w') as tar:
                    tar.add(export_folder, arcname=os.path.basename(export_folder))                                   
                             
        except Exception as e:
            await error_dialog(
                message=gettext("Unable to save model “{filename}”.").format(
                    filename=filename
                ),
                secondary_message=error_message(e),
                window=self.parent_window,
            )
            raise
        else:
            self.filename = filename
        finally:
            if status_window:
                status_window.done()
            self.event_manager.handle(Notification(gettext("Model has been saved.")))

    @property
    def parent_window(self):
        return self.main_window.window if self.main_window else None

    @action(name="file-save", shortcut="<Primary>s")
    async def action_save(self):
        """Save the file. Depending on if there is a file name, either perform
        the save directly or present the user with a save dialog box.

        Returns True if the saving actually succeeded.
        """

        if filename := self.filename:
            await self.save(filename)
        else:
            await self.action_save_as()

    @action(name="file-save-as", shortcut="<Primary><Shift>s")
    async def action_save_as(self):
        """Save the model in the element_factory by allowing the user to select
        a file name."""

        if self.filename.suffix == DEFAULT_EXT:
            filename = await save_file_dialog(
                gettext("Save Gaphor Model As"),
                self.filename or Path(gettext("New Model")).with_suffix(DEFAULT_EXT),
                parent=self.parent_window,
                filters=[(gettext("Gaphor Models"), "*.acsem", "application/x-gaphor")]
            )
        else:
            filename = await save_file_dialog(
                gettext("Save Gaphor Model As"),
                self.filename or Path(gettext("New Model")).with_suffix(".gaphor"),
                parent=self.parent_window,
                filters=[(gettext("Gaphor Models"), "*.gaphor", "application/x-gaphor")],
            )
        await self.save(filename)
        
    #Our Save Method
    @action(name="file-save-as-word")
    async def action_save_as_word(self):
        if self.filename.suffix != DEFAULT_EXT:
            await error_dialog(
                message=gettext("不支持的保存类型"),
                secondary_message=gettext("只支持从.acsem文件保存为Word文档。"),
                window=self.parent_window,
            )
            return
        
        src_filename = self._filename_realread.with_suffix(DOC_EXT)
         # 查找是否存在我们的模型文件
        if not src_filename:
            await error_dialog(
                message=gettext("无法导出模型"),
                secondary_message=gettext("模型文件被损坏，无法找到导出所需的模型数据。"),
                window=self.parent_window,
            )
            return # 目录下没找到我们的模型文件，无法导出
        
        word_FILTER = [(gettext("Word Document"), "*.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
        default_name = (self.filename or Path(gettext("Word Document"))).with_suffix(".docx")
        # 选择保存路径
        save_filename = await save_file_dialog(
            gettext("导出为Word"),
            default_name,
            parent=self.parent_window,
            filters=word_FILTER,
        )
        
        if not save_filename:
            return  # 取消

        # 3） 执行导出
        shutil.copy2(src_filename, save_filename)
        
    @action(name="file-format-document")
    async def action_format_document(self):
        word_FILTER = [(gettext("Word Document"), "*.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
        default_name = (self.filename or Path(gettext("Word Document"))).with_suffix(".docx")
        dir_path_str = str(default_name.parent)

        #选择文档
        open_path = await open_file_dialog(
            gettext("选择要格式标准化的文档"),
            self.parent_window,
            dirname=dir_path_str if self.filename else None,
            filters=word_FILTER,
            multiple=False
        )

        if not open_path:
            return  # 用户取消
        
        if not open_path.suffix.lower() == ".docx":
            await error_dialog(
                message=gettext("无效的文件类型"),
                secondary_message=gettext("请选择一个有效的Word文档（.docx）。"),
                window=self.parent_window,
            )
            return  # 非.docx文件，取消
        
        print(open_path)
        print(open_path.stem + "_formatted.docx")
        save_path = await save_file_dialog(
            gettext("保存格式标准化后的文档"),
            Path(open_path.parent, open_path.stem + "_formatted.docx"),
            parent=self.parent_window,
            filters=word_FILTER,
        )

        if not save_path:
            return  # 用户取消

        try:
            # TODO: 用你的实现替换这行：
            utils.format.docx_format(open_path, save_path, files("gaphor.ui.utils").joinpath("template.docx"), files("gaphor.ui.utils").joinpath("style_map.json"))
            # 例如：await self._export_model_to_word(src_path, dst_path)
        except Exception as e:
            await error_dialog(
                message=gettext("格式标准化失败"),
                secondary_message=str(e),
                window=self.parent_window,
            )
    
    @action(name="open-acsemai")
    async def action_open_acsemai(self):
        title = "AcsemAI智能搜索模型"
        win = AcsemAIWindow(
            parent_window=self.main_window.window,
            title=title
        )
        win.present()
        
    @event_handler(SessionCreated)
    async def _on_session_created(self, event: SessionCreated) -> None:
        if event.filename:
            await self.load(event.filename)
            self.event_manager.handle(ModelReady(self))
        elif event.template:
            self.load_template(event.template)
        else:
            self.event_manager.handle(ModelReady(self, filename=event.filename))

    @event_handler(SessionShutdownRequested)
    async def _on_session_shutdown_request(
        self, _event: SessionShutdownRequested
    ) -> None:
        """Ask user to close window if the model has changed.

        The user is asked to either discard the changes, keep the
        application running or save the model and quit afterwards.
        """

        def confirm_shutdown():
            self.event_manager.handle(SessionShutdown())

        if self.main_window.model_changed:
            answer = await save_changes_before_close_dialog(self.parent_window)
            if answer == "save":
                if filename := self.filename:
                    await self.save(filename)
                    confirm_shutdown()
                else:
                    filename = await save_file_dialog(
                        gettext("Save Gaphor Model As"),
                        self.filename
                        or Path(gettext("New Model")).with_suffix(".gaphor"),
                        parent=self.parent_window,
                        filters=GAPHOR_FILTER,
                    )
                    if filename:
                        await self.save(filename)
                        confirm_shutdown()
            elif answer == "discard":
                confirm_shutdown()
        else:
            confirm_shutdown()


async def resolve_merge_conflict_dialog(window: Gtk.Window) -> str:
    dialog = Adw.AlertDialog.new(
        gettext("Resolve Merge Conflict?"),
        gettext(
            "The model you are opening contains a merge conflict. Do you want to open the current model or the incoming change to the model?"
        ),
    )
    dialog.add_response("cancel", gettext("Cancel"))
    dialog.add_response("manual", gettext("Open Merge Editor"))
    dialog.add_response("current", gettext("Open Current"))
    dialog.add_response("incoming", gettext("Open Incoming"))
    dialog.set_close_response("cancel")

    return str(await dialog.choose(window))


async def save_changes_before_close_dialog(window: Gtk.Window) -> str:
    title = gettext("Save Changes?")
    body = gettext(
        "The open model contains unsaved changes. Changes which are not saved will be permanently lost."
    )
    dialog = Adw.AlertDialog.new(title, body)
    dialog.add_response("cancel", gettext("Cancel"))
    dialog.add_response("discard", gettext("Discard"))
    dialog.add_response("save", gettext("Save"))
    dialog.set_response_appearance("discard", Adw.ResponseAppearance.DESTRUCTIVE)
    dialog.set_response_appearance("save", Adw.ResponseAppearance.SUGGESTED)
    dialog.set_default_response("save")
    dialog.set_close_response("cancel")

    window.present()

    return str(await dialog.choose(window))

class AcsemAIWindow(Gtk.ApplicationWindow):
    def __init__(self, parent_window: Gtk.Window, title):
        app = parent_window.get_application() if parent_window else None

        super().__init__(application=app, title=title)
        if parent_window:
            self.set_transient_for(parent_window)

        self.set_default_size(440, 275)
        self.parent_window = parent_window

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
        self.input_scrolled.set_size_request(350, 180)
        
        # 执行按钮
        self.run_btn_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        self.run_btn_box.set_homogeneous(True)
        run_btn = Gtk.Button(label="搜索")
        run_btn.add_css_class("search-run")  # 命中上面的 button.search-run
        run_btn.connect("clicked", self._on_run_clicked)
        chat_btn = Gtk.Button(label="提问")
        chat_btn.add_css_class("search-run")
        chat_btn.connect("clicked", self._on_chat_clicked)
        
        key = Gtk.EventControllerKey()
        key.connect("key-pressed", self._on_input_key_pressed)
        self.input_view.add_controller(key)

        self.run_btn_box.append(run_btn)
        self.run_btn_box.append(chat_btn)
        controls.append(self.input_scrolled)
        controls.append(self.run_btn_box)
        root.append(controls)
        
    # 读取输入文本的小工具
    def _get_input_text(self) -> str:
        buf = self.input_view.get_buffer()
        start, end = buf.get_bounds()
        return buf.get_text(start, end, True)

    # 点击“执行”后的行为（你在这里写调用逻辑）
    def _on_run_clicked(self, _button):
        _, _ = self.run_action()
        
    def run_action(self):
        text = self._get_input_text().strip()
        self.text = text
        # 向大模型查询结果
        image_results, table_results = utils.Acsemai.get_research(text)
        
        image_win_title = f"智能检索到相关图片：共{len(image_results)}项"
        table_win_title = f"智能检索到相关表格：共{len(table_results)}项"
        
        # 根据text对内容做mask
        searcch_image, search_table = self._maskSearchResult(text)
        
        if searcch_image:
            win_image = ImagesViewerWindow(
                parent_window=self.parent_window,
                datalist=image_results,
                title=image_win_title
            )
            win_image.present()
        
        if search_table:
            win_table = TablesViewerWindow(
                parent_window=self.parent_window,
                datalist=table_results,
                title=table_win_title
            )
            win_table.present()
            
        print("[AcsemAIWindow] run with:", text)
        return image_results, table_results
        
    def _on_chat_clicked(self, _button):
        image_results, table_results = self.run_action()
        #弹出相关搜索窗口后，准备问答界面
        searcch_image, search_table = self._maskSearchResult(self.text)
        if not searcch_image:
            image_results=[]
        if not search_table:
            table_results=[]
            
        win = AcsemAIAllChatWindow(parent_window=self, title="AcsemAI 智能问答模型", image_results=image_results, table_results=table_results)
        win.present()
        
    # 回车执行
    def _on_input_key_pressed(self, controller, keyval, keycode, state):
        if keyval in (Gdk.KEY_Return, Gdk.KEY_KP_Enter):
            self._on_run_clicked(None)
            # ✅ 返回 True 表示我们处理了这个按键，
            # TextView 不应该继续接收这一回车（避免换行）
            return True
        return False
    
    @staticmethod
    def _maskSearchResult(text):
        searcch_image = True
        search_table = True
        if "表" in text and "图" not in text:
            searcch_image = False
        elif "图" in text and "表" not in text:
            search_table = False
        return searcch_image, search_table
    

        

