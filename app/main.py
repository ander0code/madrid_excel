from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from api import marcaciones
from config import settings

app = FastAPI(
    title=settings.APP_TITLE,
    version=settings.APP_VERSION,
    description=settings.APP_DESCRIPTION
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def read_root():
    return {"message": "API de Control de Marcaciones. Accede a /docs para más información."}


app.include_router(marcaciones.router, prefix="/api", tags=["marcaciones"])
