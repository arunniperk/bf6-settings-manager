"""Hardware detection and per-setup settings recommendation."""

import logging
import platform
import re

logger = logging.getLogger(__name__)

# Tier keys, ordered worst to best
GPU_TIERS = ["entry", "mid", "high", "enthusiast"]
CPU_TIERS = ["low", "mid", "high"]
RAM_TIERS = ["8", "16", "32"]

# GPU family patterns per tier (checked best tier first)
_GPU_PATTERNS = {
    "enthusiast": [
        r"RTX\s*50[89]0", r"RTX\s*409[00]", r"RTX\s*4080",
        r"RX\s*9070\s*XT", r"RX\s*79[00]0\s*XT", r"RX\s*7900",
    ],
    "high": [
        r"RTX\s*507[00]", r"RTX\s*407[00]", r"RTX\s*30[89]0",
        r"RX\s*9070", r"RX\s*9060\s*XT", r"RX\s*78[00]0", r"RX\s*7900\s*GRE",
        r"RX\s*68[00]0", r"RX\s*69[05]0",
    ],
    "mid": [
        r"RTX\s*[45]06[00]", r"RTX\s*30[67]0", r"RTX\s*20[678]0",
        r"RX\s*7[67]0[00]", r"RX\s*66[05]0", r"RX\s*67[05]0", r"RX\s*570[00]",
        r"Arc\s*[AB]\d{3}",
    ],
}

# Anything matching these is integrated graphics regardless of other hits
_IGPU_PATTERNS = [
    r"Radeon\(TM\)\s*Graphics", r"Radeon\s*Graphics$", r"Vega\s*\d+\s*Graphics",
    r"Intel.*(UHD|Iris|HD)\s*Graphics", r"Microsoft Basic",
]


def classify_gpu(name: str) -> str:
    """Classify a GPU name into a tier."""
    if not name:
        return "entry"
    for pattern in _IGPU_PATTERNS:
        if re.search(pattern, name, re.IGNORECASE):
            return "entry"
    for tier in ("enthusiast", "high", "mid"):
        for pattern in _GPU_PATTERNS[tier]:
            if re.search(pattern, name, re.IGNORECASE):
                return tier
    return "entry"


def classify_cpu(name: str, physical_cores: int) -> str:
    """Classify a CPU into a tier by core count and family."""
    cores = physical_cores or 0
    strong_family = bool(re.search(
        r"Ryzen\s*[79]|i[79]-|Core\s*Ultra\s*[79]", name or "", re.IGNORECASE
    ))
    if cores >= 8 or (cores >= 6 and strong_family):
        return "high"
    if cores >= 6:
        return "mid"
    return "low"


def classify_ram(total_gb: float) -> str:
    """Classify system RAM into a tier."""
    if total_gb >= 24:
        return "32"
    if total_gb >= 12:
        return "16"
    return "8"


def _detect_cpu_name() -> str:
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"HARDWARE\DESCRIPTION\System\CentralProcessor\0",
        )
        return winreg.QueryValueEx(key, "ProcessorNameString")[0].strip()
    except Exception as e:
        logger.warning(f"CPU name detection failed: {e}")
        return platform.processor() or "Unknown CPU"


def _detect_gpu_names() -> list:
    try:
        import win32com.client
        wmi = win32com.client.GetObject("winmgmts:")
        return [
            gpu.Name for gpu in wmi.InstancesOf("Win32_VideoController")
            if getattr(gpu, "Name", None)
        ]
    except Exception as e:
        logger.warning(f"GPU detection failed: {e}")
        return []


def _detect_ram_gb() -> float:
    try:
        import psutil
        return psutil.virtual_memory().total / (1024 ** 3)
    except Exception as e:
        logger.warning(f"RAM detection failed: {e}")
        return 0.0


def _detect_physical_cores() -> int:
    try:
        import psutil
        return psutil.cpu_count(logical=False) or 0
    except Exception:
        return 0


def detect_hardware() -> dict:
    """Detect CPU, GPU and RAM; return names and classified tiers.

    Blocking (registry/WMI) - call via asyncio.to_thread from async code.
    """
    cpu_name = _detect_cpu_name()
    cores = _detect_physical_cores()
    gpu_names = _detect_gpu_names()
    ram_gb = _detect_ram_gb()

    # With multiple adapters (iGPU + dGPU) use the best one
    best_gpu_name, best_gpu_tier = "Unknown GPU", "entry"
    for name in gpu_names:
        tier = classify_gpu(name)
        if GPU_TIERS.index(tier) >= GPU_TIERS.index(best_gpu_tier):
            best_gpu_name, best_gpu_tier = name, tier

    return {
        "cpu_name": cpu_name,
        "cpu_cores": cores,
        "cpu_tier": classify_cpu(cpu_name, cores),
        "gpu_name": best_gpu_name,
        "gpu_tier": best_gpu_tier,
        "ram_gb": round(ram_gb),
        "ram_tier": classify_ram(ram_gb),
    }


def recommend(gpu_tier: str, cpu_tier: str, ram_tier: str) -> dict:
    """Compute recommended settings values for a hardware combination.

    Returns {"graphics": {setting_id: "0".."3"}, "resolution_scale": float}.
    Terrain/Undergrowth stay low in every profile: short grass makes prone
    enemies easier to spot, so raising them is a competitive downgrade.
    """
    profiles = {
        "enthusiast": {
            "texture_quality": "3", "mesh_quality": "3", "lighting_quality": "3",
            "shadow_quality": "3", "effects_quality": "3", "postprocess_quality": "3",
            "terrain_quality": "1", "undergrowth_quality": "0",
            "volumetric_quality": "3", "reflection_quality": "3", "ambient_occlusion": "3",
        },
        "high": {
            "texture_quality": "3", "mesh_quality": "3", "lighting_quality": "2",
            "shadow_quality": "2", "effects_quality": "2", "postprocess_quality": "2",
            "terrain_quality": "1", "undergrowth_quality": "0",
            "volumetric_quality": "2", "reflection_quality": "2", "ambient_occlusion": "2",
        },
        "mid": {
            "texture_quality": "2", "mesh_quality": "2", "lighting_quality": "1",
            "shadow_quality": "1", "effects_quality": "1", "postprocess_quality": "1",
            "terrain_quality": "1", "undergrowth_quality": "0",
            "volumetric_quality": "0", "reflection_quality": "1", "ambient_occlusion": "1",
        },
        "entry": {
            "texture_quality": "1", "mesh_quality": "1", "lighting_quality": "0",
            "shadow_quality": "0", "effects_quality": "0", "postprocess_quality": "0",
            "terrain_quality": "0", "undergrowth_quality": "0",
            "volumetric_quality": "0", "reflection_quality": "0", "ambient_occlusion": "0",
        },
    }
    graphics = dict(profiles.get(gpu_tier, profiles["entry"]))

    # Weak CPU: effects/undergrowth are the CPU-heavy settings
    if cpu_tier == "low":
        graphics["effects_quality"] = str(min(int(graphics["effects_quality"]), 1))
        graphics["undergrowth_quality"] = "0"

    # Low RAM: texture streaming suffers first
    if ram_tier == "8":
        graphics["texture_quality"] = str(min(int(graphics["texture_quality"]), 1))

    resolution_scale = 0.75 if gpu_tier == "entry" else 1.0

    return {"graphics": graphics, "resolution_scale": resolution_scale}
