# 🎨 PPTX Generator

AI-powered presentation generator using HuggingFace models.

- **Text**: Mistral 7B for slide content generation
- **Images**: FLUX.1 Schnell for slide illustrations
- **Output**: .pptx with dark theme, images, and structured content

## Setup

```bash
pip install python-pptx requests Pillow
```

## Usage

```bash
# With API key as argument
python generate.py "Artificial Intelligence in 2026" -k hf_YOUR_KEY

# With env variable
export HF_API_KEY=hf_YOUR_KEY
python generate.py "Blockchain Technology" -n 8 -o blockchain.pptx

# Without images (faster)
python generate.py "Web Development" --no-images
```

## Options

| Flag | Description | Default |
|------|-------------|---------|
| `-o` | Output file path | `presentation.pptx` |
| `-n` | Number of slides | 6 |
| `-k` | HuggingFace API key | `$HF_API_KEY` |
| `--no-images` | Skip image generation | false |
