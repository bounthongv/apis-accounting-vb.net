from fastapi import FastAPI
from .routers import users
from .routers import chart_of_accounts
from .routers import journal_entries
from .routers import parties
from .routers import assets
from .routers import account_types
from .routers import account_groups
from .routers import transaction_types
from .routers import financial_years
from .routers import currencies
from .auth import routes as auth_routes

app = FastAPI(title="Accounting API", version="1.0.0")

# Include routers
app.include_router(auth_routes.router)
app.include_router(users.router)
app.include_router(chart_of_accounts.router)
app.include_router(journal_entries.router)
app.include_router(parties.router)
app.include_router(assets.router)
app.include_router(account_types.router)
app.include_router(account_groups.router)
app.include_router(transaction_types.router)
app.include_router(financial_years.router)
app.include_router(currencies.router)

@app.get("/")
def read_root():
    return {"message": "Welcome to the Accounting API"}

@app.get("/health")
def health_check():
    return {"status": "healthy"}