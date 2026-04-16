# app_builder.py
from pathlib import Path
import sys




app_path = Path(sys.executable).resolve() if getattr(sys, "frozen", False) else Path(__file__).resolve()
start_dir = app_path.parent
class Folders():
    def __init__(self, root: Path):
        # root
        self.app =       root / "app"
        self.core =      self.app / "core"
        self.gui =       self.app / "gui"
        self.io =        self.app / "io"
        self.models =    self.app / "models"
        self.workspace = self.app / "workspace"

        # app/workspace
        self.input_collected =   self.workspace / "input_collected"
        self.intermediate =      self.workspace / "intermediate"
        self.output =            self.workspace / "output"
        self.reports =           self.workspace / "reports"
        self.config =            self.workspace / "config"

    def getPathAttrs(self, mode="name"):
        pathAttrs = {}
        for attr, path in self.__dict__.items():
            pathAttrs[attr] = path
        if mode == "name":
            return list(pathAttrs.keys())
        if mode == "path":
            return list(pathAttrs.values())
        if mode == "all":
            return pathAttrs
        raise ValueError (f"Неподдерживаемый {mode}, Допустимые: name, path, all")
    
    def build_app(self):
        for path in self.getPathAttrs("path"):
            path.mkdir(parents = True, exist_ok = True)
    
proj_folders = Folders(start_dir)
proj_folders.build_app()


    









# {"app": start_dir{"core": start, "gui":, "io":, "models":, "workspace":("input_collected")}} #workspace должен быть указан пользователем, ибо output неудобно смотреть в папке приложения