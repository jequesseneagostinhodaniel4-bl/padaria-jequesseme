import os
import uvicorn
from app import app
if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    print(f"Padaria a correr em http://127.0.0.1:{port}")
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=True)