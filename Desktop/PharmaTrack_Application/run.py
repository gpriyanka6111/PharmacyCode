# Entry point: constructs the Flask app and launches it via FlaskUI for desktop use.

import importlib

from flaskwebgui import FlaskUI

_app_module = importlib.import_module('app$')
app = _app_module.app

if __name__ == '__main__':
    # Use FlaskWebGUI to create a desktop window
    FlaskUI(
        app=app,
        server="flask",
        width=1200,
        height=800,
        port=5000
    ).run()