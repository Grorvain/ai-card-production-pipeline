# ai-card-production-pipeline
End-to-end automated card game asset generation (ComfyUI + Ollama VLM + BiRefNet) and 300 DPI Python DTP assembly engine optimized for 6GB VRAM.

# Automated Card Game Asset & DTP Production Pipeline

An end-to-end, locally hosted production pipeline built to generate, isolate, composite, and typeset a 27-card custom deck at print-ready **300 DPI** resolution—engineered specifically to run smoothly on consumer hardware with **6GB VRAM**.

![Workflow Architecture](docs/workflow_overview.png)

## System Architecture

The pipeline is decoupled into three modular stages to maximize hardware efficiency and allow rapid single-asset hotfixes without re-running entire batches:

### 1. Hybrid VLM-to-Diffusion Workflow (`ComfyUI` + `Ollama`)
* **Switchable Routing (`CR Image Input Switch`):** Seamlessly toggles between single-image hotfix mode and full directory batch processing.
* **Automated Vision-Language Captioning:** Uses local `qwen3-vl:latest` via Ollama to analyze wildlife reference photos and extract structured anatomical/lighting prompts.
* **6GB VRAM Memory Management:** Enforces `keep_alive: 0` on the VLM node to immediately offload the language model from VRAM before diffusion sampling begins, paired with a **Single-Loader Pattern** (sharing UNET, CLIP, and VAE across tracks).
* **Dual-Track Generation:**
  * *Foreground Track:* Locks character consistency and Korean Manhwa linework using fixed seeds, stacked LoRAs, and automated background removal via **BiRefNet**.
  * *Background Track:* Generates varied atmospheric environments (rain, mist, moonlit skies) using randomized prompt pools and unlocked seeds before masked compositing.

### 2. Standalone Alpha Cutout Utility (`ComfyUI`)
* A lightweight, decoupled graph (`BiRefNet` -> `InvertMask` -> `Join Image with Alpha`) for rapid batch extraction of transparent PNG layers without loading heavy diffusion checkpoints into memory.

### 3. Automated Python DTP Assembly Engine (`scripts/build_deck.py`)
* **Data-Driven Layout:** Reads card names (Polish + Latin taxonomy), stats (HP/Attack), and ability descriptions directly from `.xlsx` / `.csv` spreadsheets via `openpyxl`.
* **Print-Ready Output (300 DPI):** Automatically wraps and centers typography with dynamic contrast strokes, aligns UI icons, composites card frames (`80x120 mm` / `945x1417 px`), and exports:
  * Individual high-res PNG cards (`output_cards/`)
  * Complete print-ready PDF deck (`talia_nietoperzy_druk.pdf`)
  * 1:1 A4 test sheets with crop marks for physical print verification.

## Tech Stack
* **AI & Vision:** ComfyUI, Ollama (`qwen3-vl`), BiRefNet, Stable Diffusion / LoRA
* **Scripting & DTP:** Python 3, Pillow (PIL), `openpyxl`
* **Hardware Target:** NVIDIA GPU (6GB VRAM) / 16GB System RAM
