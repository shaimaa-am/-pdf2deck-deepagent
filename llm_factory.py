import os

# What differs between providers: LangChain's name for it, the environment variable
# that holds its API key (None = no key), and the pip package it needs.
PROVIDERS = {
    "ollama": {"label": "Ollama (local)", "lc": "ollama", "env_key": None, "pkg": "langchain-ollama"},
    "anthropic": {"label": "Anthropic (Claude)", "lc": "anthropic", "env_key": "ANTHROPIC_API_KEY", "pkg": "langchain-anthropic"},
    "google": {"label": "Google (Gemini)", "lc": "google_genai", "env_key": "GOOGLE_API_KEY", "pkg": "langchain-google-genai"},
    "openai": {"label": "OpenAI", "lc": "openai", "env_key": "OPENAI_API_KEY", "pkg": "langchain-openai"},
    "groq": {"label": "Groq", "lc": "groq", "env_key": "GROQ_API_KEY", "pkg": "langchain-groq"},
}
def mask_key(key):
    key = (key or "").strip()
    return "" if not key else (key[:3] + "..." + key[-2:] if len(key) > 8 else "***")



def build_chat_model(provider, model, api_key=None):
    """No defaults: the user must choose the provider and the model."""
    provider = (provider or "").strip().lower()
    if not provider:
        raise ValueError("Choose a provider: " + ", ".join(PROVIDERS))
    if provider not in PROVIDERS:
        raise ValueError(f"Unknown provider '{provider}'. Choose one of: " + ", ".join(PROVIDERS))
    spec = PROVIDERS[provider]

    model = (model or "").strip()
    if not model:
        raise ValueError(f"Enter a model name for {spec['label']}")

    if spec["env_key"]:
        
        if api_key and api_key.strip():
            os.environ[spec["env_key"]] = api_key.strip()
        if not os.environ.get(spec["env_key"]):
            raise ValueError(f"Missing API key for {spec['label']}: enter it in the UI or set {spec['env_key']} in .env")

    if provider == "ollama":
        from langchain_ollama import ChatOllama
        return ChatOllama(model=model, temperature=0)

    from langchain.chat_models import init_chat_model
    try:
        return init_chat_model(model=model, model_provider=spec["lc"], api_key=api_key)
    except ImportError as e:
        raise ImportError(f"Missing package for {spec['label']}. Run: pip install {spec['pkg']}") from e
