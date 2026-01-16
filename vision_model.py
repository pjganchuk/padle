"""
Vision Model Module for PADLE

Local Moondream inference using Hugging Face Transformers.
Eliminates the need for moondream-station server.

Cross-platform compatible (Windows, macOS, Linux)
Supports CUDA, MPS (Apple Silicon), and CPU inference.
"""

import os
import sys
import logging
from typing import Optional, Callable, Dict, Any
from PIL import Image

logger = logging.getLogger(__name__)


def _detect_device() -> str:
    """
    Auto-detect the best available device for inference.
    
    Returns:
        Device string: "cuda", "mps", or "cpu"
    """
    import torch
    
    if torch.cuda.is_available():
        logger.info("CUDA GPU detected")
        return "cuda"
    
    if hasattr(torch.backends, 'mps') and torch.backends.mps.is_available():
        logger.info("Apple Silicon MPS detected")
        return "mps"
    
    logger.info("No GPU detected, using CPU")
    return "cpu"


def _get_dtype(device: str):
    """
    Get the appropriate dtype for the device.
    
    CPU requires float32, GPU can use float16/bfloat16.
    """
    import torch
    
    if device == "cpu":
        return torch.float32
    elif device == "mps":
        return torch.float16
    else:
        return torch.float16


class MoondreamLocal:
    """
    Local Moondream vision-language model wrapper.
    
    Uses Hugging Face Transformers for inference without requiring
    a separate moondream-station server process.
    """
    
    MODEL_ID = "vikhyatk/moondream2"
    MODEL_REVISION = "2024-08-26"
    
    def __init__(self):
        """Initialize the model wrapper (model not loaded yet)."""
        self._model = None
        self._tokenizer = None
        self._device = None
        self._dtype = None
        self._is_loading = False
        self._load_error: Optional[str] = None
    
    @property
    def is_loaded(self) -> bool:
        """Check if the model is loaded and ready."""
        return self._model is not None
    
    @property
    def is_loading(self) -> bool:
        """Check if the model is currently loading."""
        return self._is_loading
    
    @property
    def device(self) -> Optional[str]:
        """Get the device the model is running on."""
        return self._device
    
    @property
    def load_error(self) -> Optional[str]:
        """Get the error message if loading failed."""
        return self._load_error
    
    def get_status(self) -> str:
        """Get a human-readable status string."""
        if self._load_error:
            return f"Error: {self._load_error}"
        if self._is_loading:
            return "Loading model..."
        if self._model is not None:
            device_name = {
                "cuda": "NVIDIA GPU",
                "mps": "Apple Silicon",
                "cpu": "CPU"
            }.get(self._device, self._device)
            return f"Ready ({device_name})"
        return "Not loaded"
    
    def load(
        self,
        device: Optional[str] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> bool:
        """
        Load the Moondream model.
        
        Args:
            device: Device to use ("cuda", "mps", "cpu", or None for auto)
            progress_callback: Optional callback for progress updates
            
        Returns:
            True if loaded successfully, False otherwise
        """
        if self._model is not None:
            return True
        
        if self._is_loading:
            return False
        
        self._is_loading = True
        self._load_error = None
        
        try:
            if progress_callback:
                progress_callback("Initializing...")
            
            import torch
            from transformers import AutoModelForCausalLM, AutoTokenizer
            
            if progress_callback:
                progress_callback("Detecting hardware...")
            
            self._device = device or _detect_device()
            self._dtype = _get_dtype(self._device)
            
            logger.info(f"Loading Moondream on {self._device} with {self._dtype}")
            
            if progress_callback:
                if self._device == "cuda":
                    progress_callback("Loading model on NVIDIA GPU...")
                elif self._device == "mps":
                    progress_callback("Loading model on Apple Silicon...")
                else:
                    progress_callback("Loading model on CPU (this may take a while)...")
            
            # Load tokenizer (required for older API)
            self._tokenizer = AutoTokenizer.from_pretrained(
                self.MODEL_ID,
                revision=self.MODEL_REVISION
            )
            
            # Load model without device_map to preserve custom Moondream class
            self._model = AutoModelForCausalLM.from_pretrained(
                self.MODEL_ID,
                revision=self.MODEL_REVISION,
                trust_remote_code=True,
            )
            
            # Move to GPU if available (after loading to preserve custom methods)
            if self._device != "cpu":
                import torch
                self._model = self._model.to(self._device)
                if self._device == "cuda":
                    self._model = self._model.half()  # Use float16 on GPU for speed
            
            if progress_callback:
                progress_callback("Model loaded successfully!")
            
            logger.info("Moondream model loaded successfully")
            return True
            
        except Exception as e:
            self._load_error = str(e)
            logger.error(f"Failed to load Moondream: {e}")
            if progress_callback:
                progress_callback(f"Error: {str(e)[:50]}")
            return False
            
        finally:
            self._is_loading = False
    
    def unload(self):
        """Unload the model to free memory."""
        if self._model is not None:
            del self._model
            self._model = None
        
        if self._tokenizer is not None:
            del self._tokenizer
            self._tokenizer = None
            
            import torch
            if torch.cuda.is_available():
                torch.cuda.empty_cache()
            
            logger.info("Moondream model unloaded")
    
    def query(
        self,
        image: Image.Image,
        prompt: str,
        settings: Optional[Dict[str, Any]] = None
    ) -> Dict[str, str]:
        """
        Ask a question about an image.
        
        Args:
            image: PIL Image to analyze
            prompt: Question or instruction
            settings: Optional dict with temperature, max_tokens, top_p
            
        Returns:
            Dict with "answer" key containing the response
            
        Raises:
            RuntimeError: If model is not loaded
        """
        if self._model is None:
            raise RuntimeError("Model not loaded. Call load() first.")
        
        # Debug: log available methods
        model_type = type(self._model).__name__
        logger.info(f"Model type: {model_type}")
        
        # Try different API versions
        if hasattr(self._model, 'answer_question'):
            # Old API (2024-08-26)
            logger.info("Using answer_question API")
            enc_image = self._model.encode_image(image)
            answer = self._model.answer_question(
                enc_image,
                prompt,
                self._tokenizer
            )
        elif hasattr(self._model, 'query'):
            # Newer API
            logger.info("Using query API")
            result = self._model.query(image, prompt)
            answer = result.get("answer", str(result))
        elif hasattr(self._model, 'generate'):
            # Fallback: use generate with processor
            logger.info("Using generate API fallback")
            raise RuntimeError(
                f"Model loaded as {model_type} without Moondream methods. "
                f"Available methods: {[m for m in dir(self._model) if not m.startswith('_')][:20]}"
            )
        else:
            available = [m for m in dir(self._model) if not m.startswith('_')][:20]
            raise RuntimeError(
                f"Unknown model API. Type: {model_type}. "
                f"Some methods: {available}"
            )
        
        return {"answer": answer}
    
    def caption(
        self,
        image: Image.Image,
        length: str = "normal",
        settings: Optional[Dict[str, Any]] = None
    ) -> Dict[str, str]:
        """
        Generate a caption for an image.
        
        Args:
            image: PIL Image to caption
            length: Caption length - "short", "normal", or "long"
            settings: Optional dict with temperature, max_tokens, top_p
            
        Returns:
            Dict with "caption" key containing the generated caption
            
        Raises:
            RuntimeError: If model is not loaded
        """
        if self._model is None:
            raise RuntimeError("Model not loaded. Call load() first.")
        
        # Use query with a caption prompt for older API
        prompt = "Describe this image."
        if length == "short":
            prompt = "Describe this image briefly."
        elif length == "long":
            prompt = "Describe this image in detail."
        
        result = self.query(image, prompt, settings)
        return {"caption": result["answer"]}


_model_instance: Optional[MoondreamLocal] = None


def get_model() -> MoondreamLocal:
    """
    Get the singleton model instance.
    
    Returns:
        The shared MoondreamLocal instance
    """
    global _model_instance
    if _model_instance is None:
        _model_instance = MoondreamLocal()
    return _model_instance


def check_model_requirements() -> tuple[bool, str]:
    """
    Check if all requirements for local model inference are met.
    
    Returns:
        Tuple of (success, message)
    """
    missing = []
    
    try:
        import torch
    except ImportError:
        missing.append("torch (PyTorch)")
    
    try:
        import transformers
    except ImportError:
        missing.append("transformers")
    
    if missing:
        return False, f"Missing packages: {', '.join(missing)}"
    
    return True, "All requirements met"


def get_model_cache_info() -> Dict[str, Any]:
    """
    Get information about the cached model.
    
    Returns:
        Dict with cache_dir, is_cached, and estimated_size
    """
    from huggingface_hub import scan_cache_dir, HfFolder
    
    cache_dir = os.environ.get(
        "HF_HOME",
        os.path.join(os.path.expanduser("~"), ".cache", "huggingface")
    )
    
    model_cached = False
    model_size = 0
    
    try:
        cache_info = scan_cache_dir(cache_dir)
        for repo in cache_info.repos:
            if "moondream" in repo.repo_id.lower():
                model_cached = True
                model_size = repo.size_on_disk
                break
    except Exception:
        pass
    
    return {
        "cache_dir": cache_dir,
        "is_cached": model_cached,
        "size_bytes": model_size,
        "size_mb": model_size / (1024 * 1024) if model_size else 0,
        "estimated_download_mb": 3740,
    }