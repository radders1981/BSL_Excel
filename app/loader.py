"""
Dynamically loads semantic_config.py from the project root and returns MODELS.
This allows the user to edit semantic_config.py without modifying app code.
"""
import importlib.util
import pathlib


def load_models():
    config_path = pathlib.Path("semantic_config.py").resolve()
    if not config_path.exists():
        raise FileNotFoundError(
            f"semantic_config.py not found at {config_path}. "
            "Run main.py from the project root directory."
        )
    spec = importlib.util.spec_from_file_location("semantic_config", config_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    if not hasattr(mod, "MODELS"):
        raise AttributeError(
            "semantic_config.py must define a top-level dict named MODELS."
        )
    return mod.MODELS
