#application.py
from app.gui.catalog_gui import CatalogPipelineGUI
from app import app_builder
if __name__ == "__main__":
    app = CatalogPipelineGUI()
    app.mainloop()