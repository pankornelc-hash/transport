import os
import secrets
import string
import smtplib
import asyncio
from datetime import datetime, timedelta, date
from typing import List, Optional
from email.message import EmailMessage
from contextlib import asynccontextmanager

from fastapi import FastAPI, Depends, HTTPException, status, BackgroundTasks
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import bcrypt
from jose import JWTError, jwt
import io

from openpyxl import Workbook

from sqlalchemy import create_engine, Column, Integer, String, Float, Boolean, ForeignKey, DateTime, Text, Date
from sqlalchemy.orm import declarative_base, sessionmaker, Session, relationship
from pydantic import BaseModel, EmailStr, ConfigDict
from enum import Enum

# ==========================================
# 1. KONFIGURACJA I BEZPIECZEŃSTWO
# ==========================================
SECRET_KEY = "super-secret-key-zarzadzanie-flota"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24 * 7  # 7 dni

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./fleet_database.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False} if "sqlite" in DATABASE_URL else {})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/token")

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ==========================================
# 2. MODELE BAZY DANYCH (SQLAlchemy)
# ==========================================
class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, unique=True, index=True, nullable=False)
    hashed_password = Column(String, nullable=False)
    role = Column(String, default="manager") 
    is_active = Column(Boolean, default=True)
    can_view_fleet = Column(Boolean, default=True)
    can_edit_fleet = Column(Boolean, default=False)

class Vehicle(Base):
    __tablename__ = "vehicles"
    id = Column(Integer, primary_key=True, index=True)
    brand = Column(String, nullable=False)
    model = Column(String, nullable=False)
    registration_number = Column(String, unique=True, index=True, nullable=False)
    driver = Column(String, nullable=True)
    
    usage_country = Column(String, nullable=True)
    company = Column(String, nullable=True)
    
    is_active = Column(Boolean, default=True)
    inactive_reason = Column(String, nullable=True)
    
    current_mileage = Column(Integer, default=0)
    
    # Gwarancja i Przegląd Tech
    warranty_end = Column(Date, nullable=True)
    inspection_end = Column(Date, nullable=True)
    inspection_reminder_days = Column(Integer, default=30) 
    
    # NOWE: UDT
    udt_end = Column(Date, nullable=True)
    udt_reminder_days = Column(Integer, default=30)
    file_udt = Column(Text, nullable=True)
    
    # NOWE: Tachograf
    tacho_end = Column(Date, nullable=True)
    tacho_reminder_days = Column(Integer, default=30)
    file_tacho = Column(Text, nullable=True)
    
    # Główne skany
    file_registration = Column(Text, nullable=True)
    file_road_card = Column(Text, nullable=True)
    file_vehicle_card = Column(Text, nullable=True) 

    services = relationship("ServiceHistory", back_populates="vehicle", cascade="all, delete-orphan")
    insurances = relationship("Insurance", back_populates="vehicle", cascade="all, delete-orphan")

class ServiceHistory(Base):
    __tablename__ = "service_history"
    id = Column(Integer, primary_key=True, index=True)
    vehicle_id = Column(Integer, ForeignKey("vehicles.id"))
    
    service_date = Column(Date, nullable=True)
    mileage = Column(Integer, nullable=True)
    description = Column(String, nullable=True)
    next_service_date = Column(Date, nullable=True)
    next_service_mileage = Column(Integer, nullable=True)
    
    vehicle = relationship("Vehicle", back_populates="services")

class Insurance(Base):
    __tablename__ = "insurances"
    id = Column(Integer, primary_key=True, index=True)
    vehicle_id = Column(Integer, ForeignKey("vehicles.id"))
    
    policy_number = Column(String, nullable=True)
    valid_from = Column(Date, nullable=True)
    valid_to = Column(Date, nullable=True)
    reminder_days = Column(Integer, default=30) 
    amount = Column(Float, nullable=True)
    is_paid = Column(Boolean, default=False)
    has_gap = Column(Boolean, default=False)
    gap_valid_to = Column(Date, nullable=True)
    file_policy = Column(Text, nullable=True) 
    
    vehicle = relationship("Vehicle", back_populates="insurances")

class SMTPConfig(Base):
    __tablename__ = "smtp_config"
    id = Column(Integer, primary_key=True, index=True)
    server = Column(String, default="smtp.gmail.com")
    port = Column(Integer, default=587)
    email = Column(String, default="")
    password = Column(String, default="")

# ==========================================
# 3. SCHEMATY PYDANTIC
# ==========================================
class UserResponse(BaseModel):
    id: int
    email: str
    role: str
    can_view_fleet: bool
    can_edit_fleet: bool
    model_config = ConfigDict(from_attributes=True)

class VehicleBase(BaseModel):
    brand: str
    model: str
    registration_number: str
    driver: Optional[str] = None
    usage_country: Optional[str] = None
    company: Optional[str] = None
    is_active: bool = True
    inactive_reason: Optional[str] = None
    warranty_end: Optional[date] = None
    inspection_end: Optional[date] = None
    inspection_reminder_days: int = 30
    current_mileage: int = 0
    
    # NOWE Pydantic
    udt_end: Optional[date] = None
    udt_reminder_days: int = 30
    tacho_end: Optional[date] = None
    tacho_reminder_days: int = 30

class VehicleCreate(VehicleBase):
    pass

class VehicleListResponse(VehicleBase):
    id: int
    model_config = ConfigDict(from_attributes=True)

class ServiceCreate(BaseModel):
    service_date: Optional[date] = None
    mileage: Optional[int] = None
    description: Optional[str] = None
    next_service_date: Optional[date] = None
    next_service_mileage: Optional[int] = None

class ServiceResponse(ServiceCreate):
    id: int
    model_config = ConfigDict(from_attributes=True)

class InsuranceCreate(BaseModel):
    policy_number: Optional[str] = None
    valid_from: Optional[date] = None
    valid_to: Optional[date] = None
    reminder_days: int = 30
    amount: Optional[float] = None
    is_paid: bool = False
    has_gap: bool = False
    gap_valid_to: Optional[date] = None

class InsuranceResponse(InsuranceCreate):
    id: int
    model_config = ConfigDict(from_attributes=True)

class FileUpload(BaseModel):
    file_type: str
    file_b64: str
    insurance_id: Optional[int] = None

class SMTPConfigModel(BaseModel):
    server: str
    port: int
    email: str
    password: str

# ==========================================
# 4. LOGIKA BIZNESOWA I AUTORYZACJA
# ==========================================
def verify_password(plain_password, hashed_password):
    return bcrypt.checkpw(plain_password.encode('utf-8'), hashed_password.encode('utf-8'))

def get_password_hash(password):
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

def create_access_token(data: dict):
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)

def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        email: str = payload.get("sub")
        if email is None: raise HTTPException(status_code=401)
    except JWTError:
        raise HTTPException(status_code=401)
    user = db.query(User).filter(User.email == email).first()
    if user is None: raise HTTPException(status_code=401)
    return user

def require_admin(current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Wymagane uprawnienia administratora")
    return current_user

def require_edit_perms(current_user: User = Depends(get_current_user)):
    if current_user.role != "admin" and not current_user.can_edit_fleet:
        raise HTTPException(status_code=403, detail="Brak uprawnień do edycji")
    return current_user

def send_email(to_email: str, subject: str, body: str):
    db = SessionLocal()
    try:
        smtp_conf = db.query(SMTPConfig).first()
        if not smtp_conf or not smtp_conf.email or not smtp_conf.password:
            return
        
        msg = EmailMessage()
        msg.set_content(body)
        msg['Subject'] = subject
        msg['From'] = f"ZARZĄDZANIE FLOTĄ <{smtp_conf.email}>"
        msg['To'] = to_email
        
        if smtp_conf.port == 465:
            with smtplib.SMTP_SSL(smtp_conf.server, smtp_conf.port) as smtp:
                smtp.login(smtp_conf.email, smtp_conf.password)
                smtp.send_message(msg)
        else:
            with smtplib.SMTP(smtp_conf.server, smtp_conf.port) as smtp:
                smtp.starttls()
                smtp.login(smtp_conf.email, smtp_conf.password)
                smtp.send_message(msg)
    except Exception as e:
        print(f"[!] Błąd wysyłania emaila: {e}")
    finally:
        db.close()

def generate_random_password(length=12):
    characters = string.ascii_letters + string.digits + "!@#$%^&*"
    return ''.join(secrets.choice(characters) for _ in range(length))

def initialize_database():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    admin_email = "kh@orbis-software.pl"
    admin = db.query(User).filter(User.email == admin_email).first()
    if not admin:
        hashed = get_password_hash("Korneliusz358@@")
        new_admin = User(email=admin_email, hashed_password=hashed, role="admin", can_view_fleet=True, can_edit_fleet=True)
        db.add(new_admin)
        db.commit()
    db.close()

async def automatic_reminder_scheduler():
    while True:
        db = SessionLocal()
        try:
            today = date.today()
            admins = db.query(User).filter(User.role == "admin").all()
            vehicles = db.query(Vehicle).filter(Vehicle.is_active == True).all()
            
            # Przeglądy, UDT, Tacho
            for v in vehicles:
                # Przegląd techniczny
                if v.inspection_end:
                    remind_date = v.inspection_end - timedelta(days=v.inspection_reminder_days)
                    if remind_date == today:
                        subject = f"Przegląd techniczny dobiega końca: {v.registration_number}"
                        body = f"Uwaga!\n\nZbliża się koniec ważności przeglądu technicznego dla pojazdu:\n\nMarka/Model: {v.brand} {v.model}\nNr Rejestracyjny: {v.registration_number}\nFirma: {v.company}\n\nData końca przeglądu: {v.inspection_end}\n\nZaloguj się do systemu ZARZĄDZANIE FLOTĄ, aby zaktualizować dane."
                        for admin in admins: send_email(admin.email, subject, body)
                
                # UDT
                if v.udt_end:
                    remind_date_udt = v.udt_end - timedelta(days=v.udt_reminder_days)
                    if remind_date_udt == today:
                        subject = f"Badanie UDT dobiega końca: {v.registration_number}"
                        body = f"Uwaga!\n\nZbliża się termin badania UDT dla pojazdu:\n\nMarka/Model: {v.brand} {v.model}\nNr Rejestracyjny: {v.registration_number}\nFirma: {v.company}\n\nTermin UDT: {v.udt_end}\n\nZaloguj się do systemu ZARZĄDZANIE FLOTĄ, aby zaktualizować dane."
                        for admin in admins: send_email(admin.email, subject, body)
                
                # Tacho
                if v.tacho_end:
                    remind_date_tacho = v.tacho_end - timedelta(days=v.tacho_reminder_days)
                    if remind_date_tacho == today:
                        subject = f"Przegląd Tachografu dobiega końca: {v.registration_number}"
                        body = f"Uwaga!\n\nZbliża się termin przeglądu (legalizacji) tachografu dla pojazdu:\n\nMarka/Model: {v.brand} {v.model}\nNr Rejestracyjny: {v.registration_number}\nFirma: {v.company}\n\nTermin Tachografu: {v.tacho_end}\n\nZaloguj się do systemu ZARZĄDZANIE FLOTĄ, aby zaktualizować dane."
                        for admin in admins: send_email(admin.email, subject, body)
            
            # Polisy
            insurances = db.query(Insurance).filter(Insurance.valid_to != None).all()
            for i in insurances:
                remind_date = i.valid_to - timedelta(days=i.reminder_days)
                if remind_date == today:
                    v = db.query(Vehicle).filter(Vehicle.id == i.vehicle_id).first()
                    if v:
                        subject = f"Polisa ubezpieczeniowa dobiega końca: {v.registration_number}"
                        body = f"Uwaga!\n\nZbliża się koniec polisy ubezpieczeniowej ({i.policy_number or 'Brak numeru'}) dla pojazdu:\n\nMarka/Model: {v.brand} {v.model}\nNr Rejestracyjny: {v.registration_number}\nFirma: {v.company}\n\nData wygaśnięcia polisy: {i.valid_to}\n\nZaloguj się do systemu, aby sprawdzić szczegóły i opłacić kolejną składkę."
                        for admin in admins:
                            send_email(admin.email, subject, body)
                        
        except Exception as e:
            pass
        finally:
            db.close()
            
        await asyncio.sleep(86400) 

@asynccontextmanager
async def lifespan(app: FastAPI):
    initialize_database()
    task = asyncio.create_task(automatic_reminder_scheduler())
    yield
    task.cancel()

app = FastAPI(title="Zarzadzanie Flota API", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# ==========================================
# 5. ENDPOINTY API
# ==========================================

@app.post("/api/token")
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.email == form_data.username).first()
    if not user or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(status_code=401, detail="Nieprawidłowy e-mail lub hasło")
    
    access_token = create_access_token(data={"sub": user.email, "role": user.role})
    return {"access_token": access_token, "token_type": "bearer", "role": user.role, "perms": {"view": user.can_view_fleet, "edit": user.can_edit_fleet}}

@app.get("/api/users", response_model=List[UserResponse])
def get_users(db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    return db.query(User).all()

@app.post("/api/users")
def create_user(user_data: dict, background_tasks: BackgroundTasks, db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    email = user_data.get("email")
    role = user_data.get("role", "manager")
    if db.query(User).filter(User.email == email).first():
        raise HTTPException(status_code=400, detail="E-mail zajęty")
    
    raw_password = generate_random_password()
    hashed = get_password_hash(raw_password)
    
    new_user = User(email=email, hashed_password=hashed, role=role, can_view_fleet=True, can_edit_fleet=(role=='admin'))
    db.add(new_user)
    db.commit()
    
    body = f"Witaj,\nUtworzono dla Ciebie konto w systemie ZARZĄDZANIE FLOTĄ.\nE-mail: {email}\nHasło: {raw_password}\nZaloguj się, aby zarządzać flotą."
    background_tasks.add_task(send_email, email, "Nowe konto ZARZĄDZANIE FLOTĄ", body)
    return {"message": "Utworzono konto i wysłano e-mail."}

@app.put("/api/users/{u_id}/permissions")
def update_user_perms(u_id: int, perms: dict, db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    u = db.query(User).filter(User.id == u_id).first()
    if u:
        u.can_view_fleet = perms.get('can_view_fleet', u.can_view_fleet)
        u.can_edit_fleet = perms.get('can_edit_fleet', u.can_edit_fleet)
        db.commit()
    return {"msg": "OK"}

@app.put("/api/users/{u_id}/password")
def update_user_password(u_id: int, payload: dict, db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    u = db.query(User).filter(User.id == u_id).first()
    if u:
        u.hashed_password = get_password_hash(payload.get('new_password'))
        db.commit()
    return {"msg": "Hasło zmienione"}

@app.delete("/api/users/{u_id}")
def delete_user(u_id: int, db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    u = db.query(User).filter(User.id == u_id).first()
    if u: db.delete(u); db.commit()
    return {"msg": "Usunięto"}

@app.get("/api/smtp-config")
def get_smtp(db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    conf = db.query(SMTPConfig).first()
    return conf if conf else {"server": "", "port": 587, "email": "", "password": ""}

@app.put("/api/smtp-config")
def save_smtp(payload: SMTPConfigModel, db: Session = Depends(get_db), current_user: User = Depends(require_admin)):
    conf = db.query(SMTPConfig).first()
    if conf:
        conf.server = payload.server
        conf.port = payload.port
        conf.email = payload.email
        conf.password = payload.password
    else:
        db.add(SMTPConfig(**payload.model_dump()))
    db.commit()
    return {"msg": "Zapisano"}

@app.post("/api/smtp-config/test")
def test_smtp(payload: SMTPConfigModel, current_user: User = Depends(require_admin)):
    try:
        if payload.port == 465:
            with smtplib.SMTP_SSL(payload.server, payload.port, timeout=5) as smtp:
                smtp.login(payload.email, payload.password)
        else:
            with smtplib.SMTP(payload.server, payload.port, timeout=5) as smtp:
                smtp.starttls()
                smtp.login(payload.email, payload.password)
        return {"msg": "Sukces!"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/api/vehicles", response_model=List[VehicleListResponse])
def get_vehicles(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    return db.query(Vehicle).order_by(Vehicle.id.desc()).all()

@app.post("/api/vehicles", response_model=VehicleListResponse)
def add_vehicle(vehicle: VehicleCreate, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    new_v = Vehicle(**vehicle.model_dump())
    db.add(new_v)
    db.commit()
    db.refresh(new_v)
    return new_v

@app.put("/api/vehicles/{v_id}", response_model=VehicleListResponse)
def update_vehicle(v_id: int, vehicle: VehicleCreate, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    v = db.query(Vehicle).filter(Vehicle.id == v_id).first()
    if not v: raise HTTPException(404)
    for key, value in vehicle.model_dump().items():
        setattr(v, key, value)
    db.commit()
    db.refresh(v)
    return v

@app.delete("/api/vehicles/{v_id}")
def delete_vehicle(v_id: int, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    v = db.query(Vehicle).filter(Vehicle.id == v_id).first()
    if not v: raise HTTPException(404)
    db.delete(v)
    db.commit()
    return {"msg": "Pojazd usunięty z bazy"}

@app.post("/api/vehicles/{v_id}/files")
def upload_file(v_id: int, payload: FileUpload, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    v = db.query(Vehicle).filter(Vehicle.id == v_id).first()
    if not v: raise HTTPException(404)
    
    if payload.file_type == 'registration':
        v.file_registration = payload.file_b64
    elif payload.file_type == 'road_card':
        v.file_road_card = payload.file_b64
    elif payload.file_type == 'vehicle_card':
        v.file_vehicle_card = payload.file_b64
    elif payload.file_type == 'udt':
        v.file_udt = payload.file_b64
    elif payload.file_type == 'tacho':
        v.file_tacho = payload.file_b64
    elif payload.file_type == 'policy' and payload.insurance_id:
        ins = db.query(Insurance).filter(Insurance.id == payload.insurance_id).first()
        if ins: ins.file_policy = payload.file_b64
        
    db.commit()
    return {"msg": "Skan poprawnie wgrany"}

@app.get("/api/vehicles/{v_id}/files/{file_type}")
def get_file(v_id: int, file_type: str, ins_id: Optional[int] = None, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if file_type == 'policy' and ins_id:
        ins = db.query(Insurance).filter(Insurance.id == ins_id).first()
        return {"file_b64": ins.file_policy if ins else None}
    
    v = db.query(Vehicle).filter(Vehicle.id == v_id).first()
    if not v: raise HTTPException(404)
    
    if file_type == 'registration': return {"file_b64": v.file_registration}
    if file_type == 'road_card': return {"file_b64": v.file_road_card}
    if file_type == 'vehicle_card': return {"file_b64": v.file_vehicle_card}
    if file_type == 'udt': return {"file_b64": v.file_udt}
    if file_type == 'tacho': return {"file_b64": v.file_tacho}
    return {"file_b64": None}

@app.get("/api/vehicles/{v_id}/services", response_model=List[ServiceResponse])
def get_services(v_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    return db.query(ServiceHistory).filter(ServiceHistory.vehicle_id == v_id).order_by(ServiceHistory.id.desc()).all()

@app.post("/api/vehicles/{v_id}/services", response_model=ServiceResponse)
def add_service(v_id: int, service: ServiceCreate, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    new_s = ServiceHistory(**service.model_dump(), vehicle_id=v_id)
    db.add(new_s)
    v = db.query(Vehicle).filter(Vehicle.id == v_id).first()
    if v and service.mileage and service.mileage > v.current_mileage:
        v.current_mileage = service.mileage
    db.commit()
    db.refresh(new_s)
    return new_s

@app.get("/api/vehicles/{v_id}/insurances", response_model=List[InsuranceResponse])
def get_insurances(v_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    return db.query(Insurance).filter(Insurance.vehicle_id == v_id).order_by(Insurance.id.desc()).all()

@app.post("/api/vehicles/{v_id}/insurances", response_model=InsuranceResponse)
def add_insurance(v_id: int, ins: InsuranceCreate, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    new_i = Insurance(**ins.model_dump(), vehicle_id=v_id)
    db.add(new_i)
    db.commit()
    db.refresh(new_i)
    return new_i

@app.put("/api/insurances/{i_id}/pay")
def pay_insurance(i_id: int, db: Session = Depends(get_db), current_user: User = Depends(require_edit_perms)):
    ins = db.query(Insurance).filter(Insurance.id == i_id).first()
    if ins: ins.is_paid = True; db.commit()
    return {"msg": "Polisa opłacona"}

@app.get("/api/reports/export")
def export_excel(db: Session = Depends(get_db)):
    vehicles = db.query(Vehicle).order_by(Vehicle.id.desc()).all()
    wb = Workbook()
    ws = wb.active
    ws.title = "Raport Floty"
    
    headers = [
        "ID", "Nr Rejestracyjny", "Marka", "Model", "Firma", "Kierowca", 
        "Kraj Użytkowania", "Przebieg (km)", "Status", "Powód Nieaktywności", 
        "Koniec Gwarancji", "Przegląd Do", "Dni Alertu Przeglądu",
        "UDT Do", "Dni Alertu UDT", "Tacho Do", "Dni Alertu Tacho",
        "Skan Dowodu", "Skan K.Drogowej", "Skan K.Pojazdu", "Skan UDT", "Skan Tacho"
    ]
    ws.append(headers)
    
    for v in vehicles:
        status_txt = "Aktywny" if v.is_active else "Nieaktywny"
        ws.append([
            v.id, v.registration_number, v.brand, v.model, v.company or "", v.driver or "",
            v.usage_country or "", v.current_mileage, status_txt, v.inactive_reason or "",
            v.warranty_end.strftime("%Y-%m-%d") if v.warranty_end else "Brak",
            v.inspection_end.strftime("%Y-%m-%d") if v.inspection_end else "Brak",
            v.inspection_reminder_days,
            v.udt_end.strftime("%Y-%m-%d") if v.udt_end else "Brak", v.udt_reminder_days,
            v.tacho_end.strftime("%Y-%m-%d") if v.tacho_end else "Brak", v.tacho_reminder_days,
            "Tak" if v.file_registration else "Nie",
            "Tak" if v.file_road_card else "Nie",
            "Tak" if v.file_vehicle_card else "Nie",
            "Tak" if v.file_udt else "Nie",
            "Tak" if v.file_tacho else "Nie"
        ])
        
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    headers_resp = {
        'Content-Disposition': 'attachment; filename="Raport_Floty_Macholl.xlsx"'
    }
    return StreamingResponse(
        stream, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        headers=headers_resp
    )

# ==========================================
# 6. FRONTEND (VUE 3 + TAILWIND)
# ==========================================
@app.get("/", response_class=HTMLResponse)
def serve_frontend():
    return """
    <!DOCTYPE html>
    <html lang="pl">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ZARZĄDZANIE FLOTĄ</title>
        <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
        <script src="https://cdn.tailwindcss.com"></script>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
            body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: #f8fafc; color: #0f172a; }
            .glass-panel { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 1rem; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); }
            .input-modern { background: #f8fafc; border: 1px solid #e2e8f0; color: #0f172a; border-radius: 0.5rem; padding: 0.6rem 1rem; font-size: 0.875rem; font-weight: 500; transition: all 0.2s ease; width: 100%; }
            .input-modern:focus { outline: none; border-color: #3b82f6; box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15); background: #ffffff; }
            .input-modern::placeholder { color: #94a3b8; }
            .input-modern:disabled { opacity: 0.6; cursor: not-allowed; }
            
            .btn-primary { background: #0f172a; color: #ffffff; font-weight: 600; padding: 0.6rem 1.25rem; border-radius: 0.5rem; transition: all 0.2s ease; border: 1px solid #0f172a; text-align: center; display: flex; align-items: center; justify-content: center; }
            .btn-primary:hover:not(:disabled) { background: #1e293b; }
            .btn-primary:active:not(:disabled) { transform: scale(0.98); }
            .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
            
            .btn-secondary { background: #ffffff; border: 1px solid #e2e8f0; color: #475569; font-weight: 600; padding: 0.6rem 1.25rem; border-radius: 0.5rem; transition: all 0.2s ease; text-align: center; }
            .btn-secondary:hover:not(:disabled) { background: #f8fafc; color: #0f172a; }
            .btn-secondary:active:not(:disabled) { transform: scale(0.98); }

            .nav-item { display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem 1rem; border-radius: 0.5rem; font-size: 0.875rem; font-weight: 600; color: #64748b; transition: all 0.2s ease; }
            .nav-item:hover { color: #0f172a; background: #f1f5f9; }
            .nav-item.active { color: #2563eb; background: #eff6ff; }

            .tab-btn { padding: 0.5rem 1rem; font-size: 0.875rem; font-weight: 600; color: #64748b; border-bottom: 2px solid transparent; transition: all 0.2s ease; }
            .tab-btn:hover { color: #0f172a; }
            .tab-btn.active { color: #2563eb; border-bottom-color: #2563eb; }

            .fade-enter-active, .fade-leave-active { transition: opacity 0.2s ease; }
            .fade-enter-from, .fade-leave-to { opacity: 0; }
            .no-scrollbar::-webkit-scrollbar { display: none; }
            .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
        </style>
    </head>
    <body class="flex flex-col md:flex-row h-screen overflow-hidden">
        <div id="app" class="h-full flex w-full relative">
            
            <transition name="fade">
                <div v-if="toast.show" class="fixed top-6 right-6 z-[100] px-5 py-3 glass-panel shadow-lg flex items-center gap-3 font-semibold text-sm border-l-4"
                     :class="toast.type === 'error' ? 'border-red-500 text-red-700' : 'border-blue-500 text-blue-700'">
                    <i class="fa-solid" :class="toast.type === 'error' ? 'fa-circle-xmark text-red-500' : 'fa-check text-blue-500'"></i>
                    {{ toast.message }}
                </div>
            </transition>

            <transition name="fade">
                <div v-if="previewModal.show" class="fixed inset-0 z-[80] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4 md:p-8">
                    <div class="glass-panel w-full max-w-4xl h-[85vh] flex flex-col overflow-hidden shadow-2xl">
                        <div class="px-5 py-4 flex justify-between items-center border-b border-slate-200">
                            <h3 class="text-base font-bold text-slate-800"><i class="fa-regular fa-file-pdf text-slate-400 mr-2"></i> Podgląd dokumentu</h3>
                            <button @click="previewModal.show = false" class="text-slate-400 hover:text-slate-700 transition"><i class="fa-solid fa-xmark text-xl"></i></button>
                        </div>
                        <div class="flex-1 bg-slate-100 p-4 md:p-6 flex items-center justify-center overflow-auto">
                            <iframe :src="previewModal.src" class="w-full h-full rounded shadow-sm border border-slate-200 bg-white"></iframe>
                        </div>
                        <div class="px-5 py-4 flex justify-end border-t border-slate-200 bg-slate-50">
                            <a :href="previewModal.src" download="skan_dokumentu" class="btn-primary flex items-center gap-2"><i class="fa-solid fa-download"></i> Pobierz plik</a>
                        </div>
                    </div>
                </div>
            </transition>

            <div v-if="!token" class="flex-1 flex flex-col items-center justify-center p-4">
                <div class="w-full max-w-sm">
                    <div class="mb-8 text-center">
                        <h2 class="text-2xl font-black text-slate-900 tracking-tight">ZARZĄDZANIE FLOTĄ</h2>
                        <p class="text-sm text-slate-500 mt-1 font-medium">Zaloguj się do systemu</p>
                    </div>
                    <div class="glass-panel p-6 sm:p-8 shadow-sm">
                        <form @submit.prevent="login" class="space-y-4">
                            <div><label class="block text-xs font-bold text-slate-600 mb-1">Adres E-mail</label><input type="email" v-model="loginData.username" required class="input-modern"></div>
                            <div><label class="block text-xs font-bold text-slate-600 mb-1">Hasło</label><input type="password" v-model="loginData.password" required class="input-modern"></div>
                            <button type="submit" class="btn-primary w-full mt-2 py-3">Zaloguj się</button>
                        </form>
                    </div>
                </div>
            </div>

            <template v-else>
                
                <div class="hidden md:flex flex-col w-64 bg-white border-r border-slate-200 z-20 h-full flex-shrink-0">
                    <div class="p-6 border-b border-slate-100">
                        <h1 class="text-lg font-black text-slate-900 tracking-tight leading-tight">ZARZĄDZANIE<br><span class="text-blue-600">FLOTĄ</span></h1>
                    </div>
                    <nav class="flex-1 px-4 py-6 space-y-1 overflow-y-auto no-scrollbar">
                        <button @click="setTab('dashboard')" class="nav-item w-full" :class="currentTab === 'dashboard' || currentTab === 'car_detail' ? 'active' : ''"><i class="fa-solid fa-layer-group w-5 text-center"></i> Baza Pojazdów</button>
                        <button v-if="user.perms?.edit" @click="setTab('add_car')" class="nav-item w-full" :class="currentTab === 'add_car' ? 'active' : ''"><i class="fa-solid fa-plus w-5 text-center"></i> Dodaj Pojazd</button>
                        <button v-if="user.role === 'admin'" @click="setTab('admin')" class="nav-item w-full" :class="currentTab === 'admin' ? 'active' : ''"><i class="fa-solid fa-sliders w-5 text-center"></i> Administracja</button>
                    </nav>
                    <div class="p-4 border-t border-slate-100 bg-slate-50">
                        <div class="flex items-center justify-between mb-3 px-2">
                            <div class="overflow-hidden">
                                <p class="text-sm font-bold text-slate-800 truncate">{{ user.email }}</p>
                                <p class="text-xs text-slate-500 font-medium capitalize">{{ user.role }}</p>
                            </div>
                        </div>
                        <button @click="logout" class="btn-secondary w-full text-sm py-2">Wyloguj</button>
                    </div>
                </div>

                <div class="md:hidden fixed top-0 w-full bg-white z-50 flex justify-between items-center p-4 border-b border-slate-200 shadow-sm">
                    <span class="font-black text-lg text-slate-900 tracking-tight">ZARZĄDZANIE <span class="text-blue-600">FLOTĄ</span></span>
                    <button @click="mobileMenuOpen = !mobileMenuOpen" class="text-slate-600 hover:text-slate-900 p-1"><i class="fa-solid fa-bars text-xl"></i></button>
                </div>
                
                <transition name="fade">
                    <div v-if="mobileMenuOpen" class="md:hidden fixed top-[61px] left-0 w-full bg-white z-40 p-4 space-y-2 border-b border-slate-200 shadow-lg">
                        <button @click="setTab('dashboard')" class="nav-item w-full" :class="currentTab === 'dashboard' ? 'active' : ''"><i class="fa-solid fa-layer-group w-5"></i> Baza Pojazdów</button>
                        <button v-if="user.perms?.edit" @click="setTab('add_car')" class="nav-item w-full" :class="currentTab === 'add_car' ? 'active' : ''"><i class="fa-solid fa-plus w-5"></i> Dodaj Pojazd</button>
                        <button v-if="user.role === 'admin'" @click="setTab('admin')" class="nav-item w-full" :class="currentTab === 'admin' ? 'active' : ''"><i class="fa-solid fa-sliders w-5"></i> Administracja</button>
                        <hr class="border-slate-100 my-2">
                        <button @click="logout" class="nav-item w-full text-red-600 hover:text-red-700 hover:bg-red-50"><i class="fa-solid fa-arrow-right-from-bracket w-5"></i> Wyloguj ({{ user.email }})</button>
                    </div>
                </transition>

                <!-- MAIN CONTENT AREA -->
                <main class="flex-1 overflow-y-auto w-full pt-20 md:pt-0 p-4 md:p-8 z-10 bg-slate-50">
                    <transition name="fade" mode="out-in">
                        
                        <!-- LISTA POJAZDÓW -->
                        <div v-if="currentTab === 'dashboard'" class="max-w-7xl mx-auto space-y-8">
                            <div class="flex flex-col md:flex-row justify-between md:items-center gap-6">
                                <div>
                                    <h1 class="text-2xl font-black text-slate-900 tracking-tight">Baza Pojazdów <span class="text-blue-600">({{ filteredVehicles.length }})</span></h1>
                                    <p class="text-sm text-slate-500 mt-1">Zarządzaj flotą i generuj raporty do analizy.</p>
                                </div>
                                <div>
                                    <button @click="downloadExcel" class="flex items-center gap-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 font-bold py-2.5 px-5 rounded-lg transition-all shadow-sm active:scale-95 w-full sm:w-auto justify-center">
                                        <i class="fa-solid fa-file-excel text-emerald-600 text-lg"></i> Eksportuj do Excela
                                    </button>
                                </div>
                            </div>
                            
                            <div class="glass-panel p-6 bg-white">
                                <h3 class="text-xs font-bold text-slate-400 uppercase tracking-widest mb-4">Opcje filtrowania</h3>
                                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 w-full">
                                    <div class="relative"><i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"></i><input type="text" v-model="filters.search" placeholder="Szukaj (rej, marka)..." class="input-modern pl-10 w-full"></div>
                                    <select v-model="filters.company" class="input-modern w-full font-medium"><option value="">Firma (Wszystkie)</option><option v-for="c in uniqueCompanies" :value="c">{{ c }}</option></select>
                                    <select v-model="filters.usage_country" class="input-modern w-full font-medium"><option value="">Kraj Użytkowania (Wszystkie)</option><option v-for="c in uniqueUseCountries" :value="c">{{ c }}</option></select>
                                    <select v-model="filters.status" class="input-modern w-full text-blue-600 font-bold"><option value="all">Status: Wszystkie</option><option value="active">Tylko Aktywne</option><option value="inactive">Tylko Nieaktywne</option></select>
                                </div>
                            </div>

                            <div class="glass-panel overflow-x-auto">
                                <table class="min-w-full text-left border-collapse">
                                    <thead class="bg-slate-50 border-b border-slate-200">
                                        <tr>
                                            <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Pojazd</th>
                                            <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Kierowca & Firma</th>
                                            <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Kraj Użytkowania</th>
                                            <th class="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Status</th>
                                            <th class="px-6 py-4"></th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-slate-100 bg-white">
                                        <tr v-for="v in filteredVehicles" :key="v.id" @click="openVehicle(v)" class="hover:bg-slate-50 cursor-pointer transition">
                                            <td class="px-6 py-5"><div class="font-bold text-slate-900">{{ v.brand }} {{ v.model }}</div><div class="text-sm font-semibold text-slate-500 mt-0.5">{{ v.registration_number }}</div></td>
                                            <td class="px-6 py-5"><div class="text-sm font-medium text-slate-900">{{ v.driver || '-' }}</div><div class="text-xs text-slate-500 mt-0.5">{{ v.company || '-' }}</div></td>
                                            <td class="px-6 py-5"><div class="text-sm text-slate-700">{{ v.usage_country || '-' }}</div></td>
                                            <td class="px-6 py-5"><span v-if="v.is_active" class="inline-flex items-center gap-1.5 px-3 py-1 rounded-md text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">Aktywny</span><span v-else class="inline-flex items-center gap-1.5 px-3 py-1 rounded-md text-xs font-semibold bg-red-50 text-red-700 border border-red-200" :title="v.inactive_reason">Nieaktywny</span></td>
                                            <td class="px-6 py-5 text-right"><button class="text-sm font-semibold text-blue-600 hover:text-blue-800 transition">Szczegóły &rarr;</button></td>
                                        </tr>
                                        <tr v-if="filteredVehicles.length === 0"><td colspan="5" class="px-6 py-12 text-center text-slate-500 font-medium text-sm">Brak pojazdów spełniających wybrane kryteria filtracji.</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- WIDOK SZCZEGÓŁOWY POJAZDU -->
                        <div v-else-if="currentTab === 'car_detail' && activeCar" class="max-w-6xl mx-auto space-y-6">
                            
                            <button @click="currentTab = 'dashboard'" class="text-sm font-semibold text-slate-500 hover:text-slate-900 flex items-center gap-2 transition w-fit"><i class="fa-solid fa-arrow-left"></i> Powrót do listy</button>
                            
                            <div class="flex flex-col sm:flex-row justify-between sm:items-end gap-4 border-b border-slate-200 pb-4">
                                <div>
                                    <div class="flex items-center gap-3 mb-2"><h1 class="text-2xl font-bold text-slate-900">{{ activeCar.brand }} {{ activeCar.model }}</h1><span class="bg-slate-100 text-slate-700 border border-slate-300 px-2.5 py-0.5 rounded text-sm font-bold uppercase tracking-wider">{{ activeCar.registration_number }}</span></div>
                                    <div class="flex items-center gap-2 text-sm"><span v-if="activeCar.is_active" class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle text-[8px] mr-1 mb-0.5"></i>Aktywny</span><span v-else class="text-red-600 font-semibold"><i class="fa-solid fa-circle text-[8px] mr-1 mb-0.5"></i>Nieaktywny: {{ activeCar.inactive_reason }}</span></div>
                                </div>
                                <div class="text-left sm:text-right bg-white px-4 py-2 rounded-lg border border-slate-200 shadow-sm">
                                    <p class="text-xs font-semibold text-slate-500 uppercase tracking-wide">Zarejestrowany Przebieg</p>
                                    <p class="text-xl font-bold text-slate-900">{{ activeCar.current_mileage || 0 }} <span class="text-sm text-slate-500 font-medium">km</span></p>
                                </div>
                            </div>

                            <div class="flex overflow-x-auto gap-4 border-b border-slate-200 no-scrollbar">
                                <button @click="subTab = 'info'" class="tab-btn whitespace-nowrap" :class="subTab === 'info' ? 'active' : ''">Informacje i Parametry</button>
                                <button @click="subTab = 'udt_tacho'" class="tab-btn whitespace-nowrap" :class="subTab === 'udt_tacho' ? 'active' : ''">UDT i Tachograf</button>
                                <button @click="subTab = 'service'" class="tab-btn whitespace-nowrap" :class="subTab === 'service' ? 'active' : ''">Historia Serwisowa</button>
                                <button @click="subTab = 'insurance'" class="tab-btn whitespace-nowrap" :class="subTab === 'insurance' ? 'active' : ''">Ubezpieczenia i Polisy</button>
                            </div>

                            <transition name="fade" mode="out-in">
                                
                                <!-- SUBTAB: INFO -->
                                <div v-if="subTab === 'info'" class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                                    <div class="glass-panel p-6 md:p-8">
                                        <h3 class="text-base font-bold text-slate-900 mb-6 border-b border-slate-100 pb-3">Parametry operacyjne</h3>
                                        <div class="space-y-4">
                                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Kraj użytkowania pojazdu</label><input type="text" v-model="activeCar.usage_country" :disabled="!user.perms?.edit" class="input-modern"></div>
                                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Firma Właściciel</label><input type="text" v-model="activeCar.company" :disabled="!user.perms?.edit" class="input-modern"></div>
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Przypisany Kierowca</label><input type="text" v-model="activeCar.driver" :disabled="!user.perms?.edit" class="input-modern"></div>
                                            </div>
                                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-4 border-t border-slate-100">
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Data końca gwarancji</label><input type="date" v-model="activeCar.warranty_end" :disabled="!user.perms?.edit" class="input-modern"></div>
                                            </div>

                                            <div class="p-4 bg-slate-50 border border-slate-200 rounded-xl mt-4">
                                                <label class="block text-sm font-bold text-slate-800 mb-3"><i class="fa-solid fa-clipboard-check text-blue-500 mr-2"></i> Przegląd Techniczny</label>
                                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                    <div><label class="block text-xs font-semibold text-slate-600 mb-1">Ważny do</label><input type="date" v-model="activeCar.inspection_end" :disabled="!user.perms?.edit" class="input-modern bg-white"></div>
                                                    <div>
                                                        <label class="block text-xs font-semibold text-slate-600 mb-1">Powiadomienie E-mail</label>
                                                        <select v-model="activeCar.inspection_reminder_days" :disabled="!user.perms?.edit" class="input-modern bg-white">
                                                            <option :value="1">1 dzień przed</option><option :value="7">7 dni przed</option><option :value="14">14 dni przed</option><option :value="30">30 dni przed</option><option :value="60">60 dni przed</option>
                                                        </select>
                                                    </div>
                                                </div>
                                            </div>
                                            
                                            <div class="pt-4 border-t border-slate-100 space-y-3">
                                                <label class="flex items-center gap-2 font-semibold text-sm text-slate-800 cursor-pointer"><input type="checkbox" v-model="activeCar.is_active" :disabled="!user.perms?.edit" class="w-4 h-4 rounded border-slate-300 accent-blue-600"> Pojazd aktywny (Używany)</label>
                                                <div v-if="!activeCar.is_active">
                                                    <label class="block text-xs font-semibold text-red-600 mb-1">Powód nieaktywności</label>
                                                    <input type="text" list="inactiveReasons" v-model="activeCar.inactive_reason" :disabled="!user.perms?.edit" placeholder="Wybierz lub wpisz..." class="input-modern border-red-200 bg-red-50 focus:border-red-500 text-red-800">
                                                    <datalist id="inactiveReasons"><option value="Sprzedany"></option><option value="Zakończenie Leasingu"></option><option value="Szkoda Całkowita"></option><option value="W naprawie"></option></datalist>
                                                </div>
                                            </div>

                                            <div class="flex flex-col sm:flex-row gap-3 pt-4">
                                                <button v-if="user.perms?.edit" @click="updateActiveCar" class="btn-primary flex-1">Zapisz informacje</button>
                                                <button v-if="user.perms?.edit" @click="deleteVehicle" class="btn-secondary flex-none text-red-600 border-red-200 hover:bg-red-50 hover:text-red-700"><i class="fa-solid fa-trash-can sm:mr-2"></i><span class="hidden sm:inline">Usuń pojazd</span></button>
                                            </div>
                                        </div>
                                    </div>

                                    <div class="glass-panel p-6 md:p-8 h-fit">
                                        <h3 class="text-base font-bold text-slate-900 mb-2 border-b border-slate-100 pb-3">Dokumenty i Skany</h3>
                                        <p class="text-xs text-slate-500 mb-6"><i class="fa-solid fa-circle-info text-blue-500 mr-1"></i> Skany polis wgrywa się w zakładce <b>Ubezpieczenia i Polisy</b>.</p>
                                        <div class="space-y-3">
                                            <div class="bg-white border border-slate-200 rounded-lg p-4 flex flex-col sm:flex-row justify-between gap-3 sm:items-center">
                                                <div><p class="font-bold text-sm text-slate-800">Dowód Rejestracyjny</p><p v-if="activeCar.file_registration" class="text-xs font-semibold text-emerald-600 mt-0.5"><i class="fa-solid fa-check mr-1"></i> Wgrany</p><p v-else class="text-xs font-medium text-red-500 mt-0.5"><i class="fa-solid fa-xmark mr-1"></i> Brak pliku</p></div>
                                                <div class="flex gap-2"><button v-if="activeCar.file_registration" @click="showPreview('registration')" class="btn-secondary text-xs py-1.5 px-3">Podgląd</button><div v-if="user.perms?.edit" class="relative"><input type="file" @change="e => uploadFileBase64(e, 'registration')" accept="image/*,application/pdf" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer"><button class="btn-primary text-xs py-1.5 px-3 w-full">Wgraj</button></div></div>
                                            </div>
                                            <div class="bg-white border border-slate-200 rounded-lg p-4 flex flex-col sm:flex-row justify-between gap-3 sm:items-center">
                                                <div><p class="font-bold text-sm text-slate-800">Karta Drogowa</p><p v-if="activeCar.file_road_card" class="text-xs font-semibold text-emerald-600 mt-0.5"><i class="fa-solid fa-check mr-1"></i> Wgrana</p><p v-else class="text-xs font-medium text-red-500 mt-0.5"><i class="fa-solid fa-xmark mr-1"></i> Brak pliku</p></div>
                                                <div class="flex gap-2"><button v-if="activeCar.file_road_card" @click="showPreview('road_card')" class="btn-secondary text-xs py-1.5 px-3">Podgląd</button><div v-if="user.perms?.edit" class="relative"><input type="file" @change="e => uploadFileBase64(e, 'road_card')" accept="image/*,application/pdf" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer"><button class="btn-primary text-xs py-1.5 px-3 w-full">Wgraj</button></div></div>
                                            </div>
                                            <div class="bg-white border border-slate-200 rounded-lg p-4 flex flex-col sm:flex-row justify-between gap-3 sm:items-center">
                                                <div><p class="font-bold text-sm text-slate-800">Karta Pojazdu</p><p v-if="activeCar.file_vehicle_card" class="text-xs font-semibold text-emerald-600 mt-0.5"><i class="fa-solid fa-check mr-1"></i> Wgrana</p><p v-else class="text-xs font-medium text-red-500 mt-0.5"><i class="fa-solid fa-xmark mr-1"></i> Brak pliku</p></div>
                                                <div class="flex gap-2"><button v-if="activeCar.file_vehicle_card" @click="showPreview('vehicle_card')" class="btn-secondary text-xs py-1.5 px-3">Podgląd</button><div v-if="user.perms?.edit" class="relative"><input type="file" @change="e => uploadFileBase64(e, 'vehicle_card')" accept="image/*,application/pdf" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer"><button class="btn-primary text-xs py-1.5 px-3 w-full">Wgraj</button></div></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- SUBTAB: UDT i TACHO -->
                                <div v-if="subTab === 'udt_tacho'" class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                                    <!-- UDT -->
                                    <div class="glass-panel p-6 md:p-8">
                                        <h3 class="text-base font-bold text-slate-900 mb-6 border-b border-slate-100 pb-3"><i class="fa-solid fa-truck-ramp-box text-blue-500 mr-2"></i> Badanie UDT</h3>
                                        <div class="space-y-4">
                                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Badanie ważne do</label><input type="date" v-model="activeCar.udt_end" :disabled="!user.perms?.edit" class="input-modern"></div>
                                                <div>
                                                    <label class="block text-xs font-semibold text-slate-600 mb-1">Powiadomienie E-mail</label>
                                                    <select v-model="activeCar.udt_reminder_days" :disabled="!user.perms?.edit" class="input-modern">
                                                        <option :value="1">1 dzień przed</option><option :value="7">7 dni przed</option><option :value="14">14 dni przed</option><option :value="30">30 dni przed</option><option :value="60">60 dni przed</option>
                                                    </select>
                                                </div>
                                            </div>
                                            <button v-if="user.perms?.edit" @click="updateActiveCar" class="btn-primary w-full mt-2">Zapisz Datę UDT</button>

                                            <div class="pt-4 border-t border-slate-100 mt-6">
                                                <div class="bg-white border border-slate-200 rounded-lg p-4 flex flex-col sm:flex-row justify-between gap-3 sm:items-center">
                                                    <div>
                                                        <p class="font-bold text-sm text-slate-800">Skan z badania UDT</p>
                                                        <p v-if="activeCar.file_udt" class="text-xs font-semibold text-emerald-600 mt-0.5"><i class="fa-solid fa-check mr-1"></i> Wgrany</p>
                                                        <p v-else class="text-xs font-medium text-red-500 mt-0.5"><i class="fa-solid fa-xmark mr-1"></i> Brak pliku</p>
                                                    </div>
                                                    <div class="flex gap-2">
                                                        <button v-if="activeCar.file_udt" @click="showPreview('udt')" class="btn-secondary text-xs py-1.5 px-3">Podgląd</button>
                                                        <div v-if="user.perms?.edit" class="relative">
                                                            <input type="file" @change="e => uploadFileBase64(e, 'udt')" accept="image/*,application/pdf" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer">
                                                            <button class="btn-primary text-xs py-1.5 px-3 w-full">Wgraj</button>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    
                                    <!-- Tachograf -->
                                    <div class="glass-panel p-6 md:p-8">
                                        <h3 class="text-base font-bold text-slate-900 mb-6 border-b border-slate-100 pb-3"><i class="fa-solid fa-stopwatch text-blue-500 mr-2"></i> Przegląd Tachografu</h3>
                                        <div class="space-y-4">
                                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Przegląd ważny do</label><input type="date" v-model="activeCar.tacho_end" :disabled="!user.perms?.edit" class="input-modern"></div>
                                                <div>
                                                    <label class="block text-xs font-semibold text-slate-600 mb-1">Powiadomienie E-mail</label>
                                                    <select v-model="activeCar.tacho_reminder_days" :disabled="!user.perms?.edit" class="input-modern">
                                                        <option :value="1">1 dzień przed</option><option :value="7">7 dni przed</option><option :value="14">14 dni przed</option><option :value="30">30 dni przed</option><option :value="60">60 dni przed</option>
                                                    </select>
                                                </div>
                                            </div>
                                            <button v-if="user.perms?.edit" @click="updateActiveCar" class="btn-primary w-full mt-2">Zapisz Datę Tacho</button>

                                            <div class="pt-4 border-t border-slate-100 mt-6">
                                                <div class="bg-white border border-slate-200 rounded-lg p-4 flex flex-col sm:flex-row justify-between gap-3 sm:items-center">
                                                    <div>
                                                        <p class="font-bold text-sm text-slate-800">Skan z przeglądu (legalizacji)</p>
                                                        <p v-if="activeCar.file_tacho" class="text-xs font-semibold text-emerald-600 mt-0.5"><i class="fa-solid fa-check mr-1"></i> Wgrany</p>
                                                        <p v-else class="text-xs font-medium text-red-500 mt-0.5"><i class="fa-solid fa-xmark mr-1"></i> Brak pliku</p>
                                                    </div>
                                                    <div class="flex gap-2">
                                                        <button v-if="activeCar.file_tacho" @click="showPreview('tacho')" class="btn-secondary text-xs py-1.5 px-3">Podgląd</button>
                                                        <div v-if="user.perms?.edit" class="relative">
                                                            <input type="file" @change="e => uploadFileBase64(e, 'tacho')" accept="image/*,application/pdf" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer">
                                                            <button class="btn-primary text-xs py-1.5 px-3 w-full">Wgraj</button>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- SUBTAB: SERWIS -->
                                <div v-if="subTab === 'service'" class="space-y-4">
                                    <div class="flex justify-between items-center bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                                        <h3 class="text-base font-bold text-slate-900">Historia Serwisowa</h3>
                                        <button v-if="user.perms?.edit" @click="modals.service = true" class="btn-primary text-xs py-2 px-4"><i class="fa-solid fa-plus mr-1.5"></i> Dodaj Wpis</button>
                                    </div>

                                    <div class="glass-panel overflow-x-auto">
                                        <table class="min-w-full text-left border-collapse">
                                            <thead class="bg-slate-50 border-b border-slate-200">
                                                <tr>
                                                    <th class="px-6 py-3 text-xs font-bold text-slate-500 uppercase tracking-wider">Data</th>
                                                    <th class="px-6 py-3 text-xs font-bold text-slate-500 uppercase tracking-wider">Przebieg</th>
                                                    <th class="px-6 py-3 text-xs font-bold text-slate-500 uppercase tracking-wider w-1/3">Zakres Prac</th>
                                                    <th class="px-6 py-3 text-xs font-bold text-slate-500 uppercase tracking-wider">Kolejny planowany</th>
                                                </tr>
                                            </thead>
                                            <tbody class="divide-y divide-slate-100 bg-white">
                                                <tr v-for="s in activeServices" :key="s.id" class="hover:bg-slate-50 transition">
                                                    <td class="px-6 py-4 text-sm font-semibold text-slate-900">{{ s.service_date || '-' }}</td>
                                                    <td class="px-6 py-4 text-sm text-slate-600">{{ s.mileage ? s.mileage + ' km' : '-' }}</td>
                                                    <td class="px-6 py-4 text-sm text-slate-700">{{ s.description || '-' }}</td>
                                                    <td class="px-6 py-4 text-sm">
                                                        <div class="font-medium text-slate-800">{{ s.next_service_date || '-' }}</div>
                                                        <div class="text-xs text-slate-500">{{ s.next_service_mileage ? s.next_service_mileage + ' km' : '' }}</div>
                                                    </td>
                                                </tr>
                                                <tr v-if="activeServices.length===0"><td colspan="4" class="px-6 py-10 text-center text-slate-500 font-medium text-sm">Brak zapisanej historii.</td></tr>
                                            </tbody>
                                        </table>
                                    </div>
                                </div>

                                <!-- SUBTAB: UBEZPIECZENIA -->
                                <div v-if="subTab === 'insurance'" class="space-y-4">
                                    <div class="flex justify-between items-center bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
                                        <h3 class="text-base font-bold text-slate-900">Dokumenty Ubezpieczeniowe</h3>
                                        <button v-if="user.perms?.edit" @click="modals.insurance = true" class="btn-primary text-xs py-2 px-4"><i class="fa-solid fa-plus mr-1.5"></i> Nowa Polisa</button>
                                    </div>

                                    <div class="grid grid-cols-1 xl:grid-cols-2 gap-6">
                                        <div v-for="ins in activeInsurances" :key="ins.id" class="glass-panel p-6 relative overflow-hidden" :class="isPastDate(ins.valid_to) ? 'border-red-300 bg-red-50/30' : ''">
                                            <div class="flex justify-between items-start mb-5 border-b border-slate-100 pb-3">
                                                <div><p class="text-xs font-semibold text-slate-500 uppercase">Numer Polisy</p><p class="text-lg font-bold text-slate-900">{{ ins.policy_number || 'Brak numeru' }}</p></div>
                                                <div v-if="ins.is_paid" class="bg-emerald-100 text-emerald-800 border border-emerald-200 px-2 py-0.5 rounded text-xs font-bold uppercase">Opłacono</div>
                                                <button v-else-if="user.perms?.edit" @click="payInsurance(ins.id)" class="btn-primary text-xs py-1 px-3 bg-rose-600 border-rose-600 hover:bg-rose-700">Zapłać</button>
                                            </div>
                                            <div class="grid grid-cols-2 gap-4">
                                                <div><p class="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Ochrona od - do</p><p class="text-sm font-semibold text-slate-800">{{ ins.valid_from || '-' }} <br> <span :class="isPastDate(ins.valid_to) ? 'text-red-600 font-bold' : ''">{{ ins.valid_to || '-' }}</span></p><p class="text-xs text-slate-500 mt-1">Alert: {{ ins.reminder_days }} dni przed</p></div>
                                                <div><p class="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Składka</p><p class="text-base font-bold text-slate-900">{{ ins.amount || 0 }} <span class="text-xs text-slate-500 font-medium">PLN</span></p><div v-if="ins.has_gap" class="mt-2 inline-flex items-center gap-1.5 bg-blue-50 border border-blue-200 text-blue-700 px-2 py-0.5 rounded text-[10px] font-bold uppercase">GAP do: {{ ins.gap_valid_to || '-' }}</div></div>
                                            </div>
                                            <div class="mt-5 pt-4 border-t border-slate-100 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                                                <div class="flex items-center gap-2"><p v-if="ins.file_policy" class="text-xs font-bold text-emerald-600"><i class="fa-solid fa-check"></i> Skan</p><p v-else class="text-xs font-bold text-red-500"><i class="fa-solid fa-xmark"></i> Skan</p></div>
                                                <div class="flex gap-2 w-full sm:w-auto"><button v-if="ins.file_policy" @click="showPreview('policy', ins.id)" class="btn-secondary text-xs py-1.5 px-3 flex-1 sm:flex-none">Podgląd</button><div v-if="user.perms?.edit" class="relative flex-1 sm:flex-none"><input type="file" @change="e => uploadFileBase64(e, 'policy', ins.id)" accept="image/*,application/pdf" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer"><button class="btn-primary text-xs py-1.5 px-3 w-full">Wgraj</button></div></div>
                                            </div>
                                        </div>
                                        <div v-if="activeInsurances.length===0" class="col-span-full glass-panel p-10 text-center text-slate-500 font-medium text-sm">Brak zarejestrowanych polis.</div>
                                    </div>
                                </div>
                            </div>

                            <!-- REJESTRACJA POJAZDU -->
                            <div v-else-if="currentTab === 'add_car'" class="max-w-4xl mx-auto space-y-6">
                                <div class="border-b border-slate-200 pb-4">
                                    <h1 class="text-2xl font-bold text-slate-900">Rejestracja Pojazdu</h1>
                                    <p class="text-sm text-slate-500 mt-1">Wprowadź nowy pojazd do systemu flotowego.</p>
                                </div>
                                <div class="glass-panel p-6 md:p-10">
                                    <form @submit.prevent="submitVehicle" class="space-y-8">
                                        <div class="space-y-4">
                                            <h4 class="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">Identyfikacja Pojazdu</h4>
                                            <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
                                                <div class="md:col-span-1"><label class="block text-xs font-semibold text-slate-600 mb-1">Nr Rejestracyjny</label><input type="text" v-model="forms.car.registration_number" required class="input-modern uppercase font-bold text-slate-900 tracking-wider"></div>
                                                <div class="md:col-span-1"><label class="block text-xs font-semibold text-slate-600 mb-1">Marka</label><input type="text" v-model="forms.car.brand" required class="input-modern"></div>
                                                <div class="md:col-span-1"><label class="block text-xs font-semibold text-slate-600 mb-1">Model</label><input type="text" v-model="forms.car.model" required class="input-modern"></div>
                                            </div>
                                        </div>
                                        <div class="space-y-4">
                                            <h4 class="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">Zarządzanie i Przypisania</h4>
                                            <div class="grid grid-cols-1 gap-5 mb-2"><div><label class="block text-xs font-semibold text-slate-600 mb-1">Kraj użytkowania pojazdu</label><input type="text" v-model="forms.car.usage_country" class="input-modern"></div></div>
                                            <div class="grid grid-cols-1 md:grid-cols-2 gap-5"><div><label class="block text-xs font-semibold text-slate-600 mb-1">Firma Właściciel</label><input type="text" v-model="forms.car.company" class="input-modern"></div><div><label class="block text-xs font-semibold text-slate-600 mb-1">Przypisany Kierowca</label><input type="text" v-model="forms.car.driver" class="input-modern"></div></div>
                                        </div>
                                        
                                        <!-- Dodano UDT i Tacho przy rejestracji -->
                                        <div class="space-y-4">
                                            <h4 class="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">Badania Dodatkowe (Opcjonalnie)</h4>
                                            <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Data ważności UDT</label><input type="date" v-model="forms.car.udt_end" class="input-modern"></div>
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Data przeglądu Tachografu</label><input type="date" v-model="forms.car.tacho_end" class="input-modern"></div>
                                            </div>
                                        </div>

                                        <div class="space-y-4">
                                            <h4 class="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2">Status i Przebieg</h4>
                                            <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Przebieg Początkowy (km)</label><input type="number" v-model="forms.car.current_mileage" class="input-modern"></div>
                                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Status Floty</label><select v-model="forms.car.is_active" class="input-modern"><option :value="true">Aktywny (Włączony do floty)</option><option :value="false">Nieaktywny</option></select></div>
                                                <div class="md:col-span-2" v-if="!forms.car.is_active"><label class="block text-xs font-semibold text-red-600 mb-1">Powód nieaktywności</label><input type="text" list="inactiveReasonsAdd" v-model="forms.car.inactive_reason" placeholder="Wybierz lub wpisz..." class="input-modern border-red-300 bg-red-50 text-red-900"><datalist id="inactiveReasonsAdd"><option value="Sprzedany"></option><option value="Zakończenie Leasingu"></option><option value="Szkoda Całkowita"></option><option value="W naprawie"></option></datalist></div>
                                            </div>
                                        </div>
                                        
                                        <div class="pt-4 border-t border-slate-100 flex justify-end">
                                            <button type="submit" class="btn-primary py-3 px-8 w-full sm:w-auto">Zarejestruj Pojazd</button>
                                        </div>
                                    </form>
                                </div>
                            </div>

                            <!-- ADMINISTRACJA -->
                            <div v-else-if="currentTab === 'admin' && user.role === 'admin'" class="space-y-6 max-w-5xl mx-auto">
                                <div class="border-b border-slate-200 pb-4">
                                    <h1 class="text-2xl font-bold text-slate-900">Administracja i Ustawienia</h1>
                                </div>
                                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                                    <div class="glass-panel p-6 sm:p-8">
                                        <h3 class="text-base font-bold text-slate-900 mb-2"><i class="fa-solid fa-satellite-dish text-blue-500 mr-2"></i> Konfiguracja Serwera SMTP</h3>
                                        <p class="text-xs text-slate-500 mb-6">Wymagane do automatycznych powiadomień E-mail.</p>
                                        <div class="space-y-4">
                                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Serwer SMTP</label><input type="text" v-model="smtpConfig.server" placeholder="smtp.gmail.com" class="input-modern w-full"></div>
                                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4"><div><label class="block text-xs font-semibold text-slate-600 mb-1">Port</label><input type="number" v-model="smtpConfig.port" class="input-modern w-full"></div><div><label class="block text-xs font-semibold text-slate-600 mb-1">E-mail nadawcy</label><input type="email" v-model="smtpConfig.email" class="input-modern w-full"></div></div>
                                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Hasło aplikacji</label><input type="password" v-model="smtpConfig.password" class="input-modern w-full font-mono"></div>
                                            <div class="flex flex-col sm:flex-row gap-3 pt-2"><button @click="saveSmtpConfig" class="btn-primary w-full sm:flex-1">Zapisz</button><button @click="testSmtpConfig" class="btn-secondary w-full sm:flex-1">Test Połączenia</button></div>
                                        </div>
                                    </div>
                                    <div class="glass-panel p-6 sm:p-8 flex flex-col">
                                        <h3 class="text-base font-bold text-slate-900 mb-2"><i class="fa-solid fa-users-gear text-blue-500 mr-2"></i> Zarządzanie Dostępem</h3>
                                        <p class="text-xs text-slate-500 mb-6">Konta uprawnione do systemu flotowego.</p>
                                        <form @submit.prevent="addUser" class="flex flex-col md:flex-row gap-3 mb-6 bg-slate-50 p-4 rounded-lg border border-slate-200">
                                            <div class="w-full md:flex-1"><input type="email" v-model="forms.user.email" placeholder="Wpisz nowy adres e-mail" required class="input-modern w-full bg-white"></div>
                                            <div class="w-full md:w-auto"><select v-model="forms.user.role" class="input-modern w-full bg-white"><option value="manager">Manager</option><option value="admin">Admin</option></select></div>
                                            <button type="submit" class="btn-primary w-full md:w-auto px-5"><i class="fa-solid fa-plus md:mr-0 mr-2"></i><span class="md:hidden"> Dodaj użytkownika</span></button>
                                        </form>
                                        <div class="overflow-x-auto rounded-lg border border-slate-200 flex-1">
                                            <table class="min-w-full text-left bg-white">
                                                <thead class="bg-slate-50 border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-widest">
                                                    <tr><th class="px-4 py-3">Użytkownik</th><th class="px-4 py-3 text-center" title="Odczyt floty"><i class="fa-regular fa-eye text-sm"></i></th><th class="px-4 py-3 text-center" title="Edycja floty"><i class="fa-solid fa-pen text-sm"></i></th><th class="px-4 py-3"></th></tr>
                                                </thead>
                                                <tbody class="divide-y divide-slate-100">
                                                    <tr v-for="u in systemUsers" :key="u.id" class="hover:bg-slate-50 transition">
                                                        <td class="px-4 py-3 whitespace-nowrap"><div class="font-semibold text-slate-900 text-sm">{{ u.email }}</div><div class="text-[10px] text-slate-500 uppercase font-bold">{{ u.role }}</div></td>
                                                        <td class="px-4 py-3 text-center"><input type="checkbox" v-model="u.can_view_fleet" @change="updateUserPerms(u)" :disabled="u.role==='admin'" class="w-4 h-4 rounded border-slate-300 accent-blue-600"></td>
                                                        <td class="px-4 py-3 text-center"><input type="checkbox" v-model="u.can_edit_fleet" @change="updateUserPerms(u)" :disabled="u.role==='admin'" class="w-4 h-4 rounded border-slate-300 accent-blue-600"></td>
                                                        <td class="px-4 py-3 text-right flex justify-end gap-1 items-center h-full pt-4"><button @click="openPasswordModal(u)" class="text-slate-400 hover:text-blue-600 transition p-1.5"><i class="fa-solid fa-key"></i></button><button @click="deleteUser(u.id)" :disabled="u.email===user.email" class="text-slate-400 hover:text-red-500 transition p-1.5 disabled:opacity-30"><i class="fa-solid fa-trash-can"></i></button></td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </transition>
                    </main>
                </div>

                <!-- MODALE WIDOKÓW -->
                <transition name="fade">
                <div v-if="modals.service" class="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4">
                    <div class="glass-panel max-w-md w-full p-6 md:p-8">
                        <h3 class="text-lg font-bold text-slate-900 mb-5 border-b border-slate-100 pb-3">Rejestracja Serwisu</h3>
                        <form @submit.prevent="addService" class="space-y-4">
                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Data wykonania (opcjonalnie)</label><input type="date" v-model="forms.service.service_date" class="input-modern"></div>
                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Przebieg (km) (opcjonalnie)</label><input type="number" v-model="forms.service.mileage" class="input-modern"></div>
                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Zakres prac (opcjonalnie)</label><textarea v-model="forms.service.description" rows="2" class="input-modern resize-none"></textarea></div>
                            <div class="grid grid-cols-2 gap-4 border-t border-slate-100 pt-4 mt-2">
                                <div><label class="block text-[10px] font-bold text-blue-600 uppercase mb-1">Planowana data</label><input type="date" v-model="forms.service.next_service_date" class="input-modern"></div>
                                <div><label class="block text-[10px] font-bold text-blue-600 uppercase mb-1">Planowany przebieg</label><input type="number" v-model="forms.service.next_service_mileage" class="input-modern"></div>
                            </div>
                            <div class="flex gap-3 pt-3"><button type="button" @click="modals.service=false" class="btn-secondary flex-1">Anuluj</button><button type="submit" class="btn-primary flex-1">Zapisz Wpis</button></div>
                        </form>
                    </div>
                </div>
                </transition>

                <transition name="fade">
                <div v-if="modals.insurance" class="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4">
                    <div class="glass-panel max-w-md w-full p-6 md:p-8 max-h-[90vh] overflow-y-auto">
                        <h3 class="text-lg font-bold text-slate-900 mb-5 border-b border-slate-100 pb-3">Wprowadzenie Polisy</h3>
                        <form @submit.prevent="addInsurance" class="space-y-4">
                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Numer Polisy (opcjonalnie)</label><input type="text" v-model="forms.insurance.policy_number" class="input-modern font-bold"></div>
                            <div class="grid grid-cols-2 gap-4">
                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Ważna od (opcjonalnie)</label><input type="date" v-model="forms.insurance.valid_from" class="input-modern"></div>
                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Ważna do (opcjonalnie)</label><input type="date" v-model="forms.insurance.valid_to" class="input-modern"></div>
                            </div>
                            <div class="bg-slate-50 p-3 border border-slate-200 rounded-lg">
                                <label class="block text-xs font-bold text-slate-700 mb-2"><i class="fa-solid fa-bell text-blue-500 mr-1"></i> Alert E-mail</label>
                                <select v-model="forms.insurance.reminder_days" class="input-modern bg-white">
                                    <option :value="1">1 dzień przed końcem</option><option :value="7">7 dni przed końcem</option><option :value="14">14 dni przed końcem</option><option :value="30">30 dni (1 miesiąc)</option><option :value="60">60 dni (2 miesiące)</option>
                                </select>
                            </div>
                            <div class="grid grid-cols-2 gap-4 border-t border-slate-100 pt-4">
                                <div><label class="block text-xs font-semibold text-slate-600 mb-1">Kwota (opcjonalnie)</label><input type="number" step="0.01" v-model="forms.insurance.amount" class="input-modern font-bold"></div>
                                <div class="flex items-end pb-2"><label class="flex items-center gap-2 text-sm font-semibold text-slate-700 cursor-pointer"><input type="checkbox" v-model="forms.insurance.is_paid" class="w-4 h-4 rounded border-slate-300 accent-blue-600"> Opłacona</label></div>
                            </div>
                            <div class="border-t border-slate-100 pt-4">
                                <label class="flex items-center gap-2 text-sm font-semibold text-slate-700 mb-3 cursor-pointer"><input type="checkbox" v-model="forms.insurance.has_gap" class="w-4 h-4 rounded border-slate-300 accent-blue-600"> Polisa zawiera GAP</label>
                                <div v-if="forms.insurance.has_gap"><label class="block text-xs font-semibold text-blue-600 mb-1">GAP ważny do daty</label><input type="date" v-model="forms.insurance.gap_valid_to" class="input-modern"></div>
                            </div>
                            <div class="flex gap-3 pt-3"><button type="button" @click="modals.insurance=false" class="btn-secondary flex-1">Anuluj</button><button type="submit" class="btn-primary flex-1">Zapisz Polisę</button></div>
                        </form>
                    </div>
                </div>
                </transition>

                <transition name="fade">
                <div v-if="modals.password" class="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4">
                    <div class="glass-panel max-w-sm w-full p-6 md:p-8">
                        <h3 class="text-lg font-bold text-slate-900 mb-1">Zmień hasło konta</h3>
                        <p class="text-xs font-medium text-slate-500 mb-5 border-b border-slate-100 pb-3">{{ forms.password.email }}</p>
                        <form @submit.prevent="submitNewPassword" class="space-y-4">
                            <div><label class="block text-xs font-semibold text-slate-600 mb-1">Nowe hasło</label><input type="text" v-model="forms.password.newPassword" required class="input-modern"></div>
                            <div class="flex gap-3 pt-2"><button type="button" @click="modals.password=false" class="btn-secondary flex-1">Anuluj</button><button type="submit" class="btn-primary flex-1">Zapisz</button></div>
                        </form>
                    </div>
                </div>
                </transition>

            </template>
        </div>

        <script>
            const { createApp } = Vue;
            createApp({
                data() {
                    return {
                        loginData: { username: '', password: '' },
                        token: localStorage.getItem('fleet_token') || null,
                        user: JSON.parse(localStorage.getItem('fleet_user')) || {},
                        currentTab: 'dashboard',
                        subTab: 'info',
                        mobileMenuOpen: false,
                        
                        toast: { show: false, message: '', type: 'success' },
                        modals: { service: false, insurance: false, password: false },
                        previewModal: { show: false, src: '' },
                        
                        vehicles: [],
                        activeCar: null,
                        activeServices: [],
                        activeInsurances: [],
                        systemUsers: [],
                        
                        smtpConfig: { server: '', port: 587, email: '', password: '' },
                        filters: { search: '', company: '', status: 'all', usage_country: '' },
                        
                        forms: {
                            car: { brand: '', model: '', registration_number: '', driver: '', usage_country: '', company: '', current_mileage: 0, is_active: true, inactive_reason: '', udt_end: '', tacho_end: '' },
                            service: { service_date: '', mileage: '', description: '', next_service_date: '', next_service_mileage: '' },
                            insurance: { policy_number: '', valid_from: '', valid_to: '', reminder_days: 30, amount: '', is_paid: false, has_gap: false, gap_valid_to: '' },
                            user: { email: '', role: 'manager' },
                            password: { userId: null, email: '', newPassword: '' }
                        }
                    }
                },
                computed: {
                    uniqueCompanies() {
                        const comps = this.vehicles.map(v => v.company).filter(c => c && c.trim() !== '');
                        return [...new Set(comps)];
                    },
                    uniqueUseCountries() {
                        const arr = this.vehicles.map(v => v.usage_country).filter(c => c && c.trim() !== '');
                        return [...new Set(arr)];
                    },
                    filteredVehicles() {
                        let filtered = this.vehicles;
                        if (this.filters.company) filtered = filtered.filter(v => v.company === this.filters.company);
                        if (this.filters.usage_country) filtered = filtered.filter(v => v.usage_country === this.filters.usage_country);
                        if (this.filters.status === 'active') filtered = filtered.filter(v => v.is_active === true);
                        else if (this.filters.status === 'inactive') filtered = filtered.filter(v => v.is_active === false);
                        if (this.filters.search) {
                            const q = this.filters.search.toLowerCase();
                            filtered = filtered.filter(v => v.registration_number.toLowerCase().includes(q) || v.brand.toLowerCase().includes(q) || v.model.toLowerCase().includes(q));
                        }
                        return filtered;
                    }
                },
                mounted() {
                    if (this.token) this.loadData();
                },
                methods: {
                    showToast(message, type = 'success') {
                        this.toast.message = message; this.toast.type = type; this.toast.show = true;
                        setTimeout(() => this.toast.show = false, 3500);
                    },
                    isPastDate(dateString) {
                        if(!dateString) return false;
                        return new Date(dateString) < new Date();
                    },
                    cleanPayload(obj) {
                        let payload = { ...obj };
                        for (let key in payload) { if (payload[key] === "") payload[key] = null; }
                        return payload;
                    },
                    async api(endpoint, method = 'GET', body = null) {
                        const headers = { 'Authorization': 'Bearer ' + this.token };
                        if (body && !(body instanceof FormData)) {
                            headers['Content-Type'] = 'application/json';
                            body = JSON.stringify(body);
                        }
                        const res = await fetch('/api/' + endpoint, { method, headers, body });
                        const data = await res.json();
                        if (!res.ok) {
                            if (res.status === 401) this.logout();
                            throw new Error(data.detail || 'Błąd API');
                        }
                        return data;
                    },
                    async login() {
                        try {
                            const fd = new FormData();
                            fd.append('username', this.loginData.username);
                            fd.append('password', this.loginData.password);
                            const res = await fetch('/api/token', { method: 'POST', body: fd });
                            const data = await res.json();
                            if (!res.ok) throw new Error(data.detail);
                            this.token = data.access_token;
                            this.user = { email: this.loginData.username, role: data.role, perms: data.perms };
                            localStorage.setItem('fleet_token', this.token);
                            localStorage.setItem('fleet_user', JSON.stringify(this.user));
                            this.loadData();
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    logout() {
                        this.token = null; this.user = {}; this.activeCar = null;
                        localStorage.removeItem('fleet_token'); localStorage.removeItem('fleet_user');
                    },
                    setTab(tab) {
                        this.currentTab = tab; this.mobileMenuOpen = false; this.activeCar = null;
                        if(tab === 'dashboard' || tab === 'admin') this.loadData();
                    },
                    async loadData() {
                        try {
                            this.vehicles = await this.api('vehicles');
                            if (this.user.role === 'admin') {
                                this.systemUsers = await this.api('users');
                                this.smtpConfig = await this.api('smtp-config');
                            }
                        } catch(e) { console.error(e); }
                    },
                    async openVehicle(v) {
                        this.activeCar = JSON.parse(JSON.stringify(v));
                        this.currentTab = 'car_detail';
                        this.subTab = 'info';
                        await this.loadVehicleDetails();
                    },
                    async loadVehicleDetails() {
                        if(!this.activeCar) return;
                        const v_id = this.activeCar.id;
                        this.activeServices = await this.api(`vehicles/${v_id}/services`);
                        this.activeInsurances = await this.api(`vehicles/${v_id}/insurances`);
                    },
                    async submitVehicle() {
                        try {
                            const payload = this.cleanPayload(this.forms.car);
                            await this.api('vehicles', 'POST', payload);
                            this.showToast("Pojazd włączony do floty");
                            this.forms.car = { brand: '', model: '', registration_number: '', driver: '', usage_country: '', company: '', current_mileage: 0, is_active: true, inactive_reason: '', udt_end: '', tacho_end: '' };
                            this.setTab('dashboard');
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    async updateActiveCar() {
                        try {
                            const payload = this.cleanPayload(this.activeCar);
                            await this.api(`vehicles/${this.activeCar.id}`, 'PUT', payload);
                            this.showToast("Zapisano ustawienia pojazdu");
                            const idx = this.vehicles.findIndex(x => x.id === this.activeCar.id);
                            if(idx !== -1) this.vehicles[idx] = {...this.activeCar};
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    async deleteVehicle() {
                        if(!confirm("CZY NA PEWNO CHCESZ TRWALE USUNĄĆ TEN POJAZD? Ta operacja usunie wszystkie przypisane skany, polisy oraz historię serwisową i jest nieodwracalna!")) return;
                        try {
                            await this.api(`vehicles/${this.activeCar.id}`, 'DELETE');
                            this.showToast("Pojazd został trwale usunięty z systemu.");
                            this.setTab('dashboard');
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    async addService() {
                        try {
                            const payload = this.cleanPayload(this.forms.service);
                            await this.api(`vehicles/${this.activeCar.id}/services`, 'POST', payload);
                            this.showToast("Zarejestrowano wpis serwisowy");
                            this.modals.service = false;
                            this.forms.service = { service_date: '', mileage: '', description: '', next_service_date: '', next_service_mileage: '' };
                            await this.loadVehicleDetails();
                            if(payload.mileage && payload.mileage > this.activeCar.current_mileage) { this.activeCar.current_mileage = payload.mileage; }
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    async addInsurance() {
                        try {
                            const payload = this.cleanPayload(this.forms.insurance);
                            await this.api(`vehicles/${this.activeCar.id}/insurances`, 'POST', payload);
                            this.showToast("Polisa dodana do systemu.");
                            this.modals.insurance = false;
                            this.forms.insurance = { policy_number: '', valid_from: '', valid_to: '', reminder_days: 30, amount: '', is_paid: false, has_gap: false, gap_valid_to: '' };
                            await this.loadVehicleDetails();
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    async payInsurance(i_id) {
                        try {
                            await this.api(`insurances/${i_id}/pay`, 'PUT');
                            this.showToast("Składka oznaczona jako opłacona");
                            await this.loadVehicleDetails();
                        } catch(e) { this.showToast(e.message, 'error'); }
                    },
                    
                    uploadFileBase64(event, fileType, insuranceId = null) {
                        const file = event.target.files[0];
                        if(!file) return;
                        if(file.size > 5 * 1024 * 1024) { this.showToast("Plik jest za duży (Max 5MB).", "error"); return; }
                        
                        const reader = new FileReader();
                        reader.onload = async (e) => {
                            const b64 = e.target.result;
                            try {
                                await this.api(`vehicles/${this.activeCar.id}/files`, 'POST', { file_type: fileType, file_b64: b64, insurance_id: insuranceId });
                                this.showToast("Dokument załadowany do bazy.");
                                if(fileType === 'registration') this.activeCar.file_registration = true;
                                if(fileType === 'road_card') this.activeCar.file_road_card = true;
                                if(fileType === 'vehicle_card') this.activeCar.file_vehicle_card = true;
                                if(fileType === 'udt') this.activeCar.file_udt = true;
                                if(fileType === 'tacho') this.activeCar.file_tacho = true;
                                if(fileType === 'policy') await this.loadVehicleDetails();
                            } catch(err) { this.showToast(err.message, 'error'); }
                        };
                        reader.readAsDataURL(file);
                    },
                    async showPreview(fileType, insuranceId = null) {
                        try {
                            let url = `vehicles/${this.activeCar.id}/files/${fileType}`;
                            if(insuranceId) url += `?ins_id=${insuranceId}`;
                            const res = await this.api(url);
                            if(!res.file_b64) { this.showToast("Brak wgranego skanu dla tego elementu", "error"); return; }
                            this.previewModal.src = res.file_b64;
                            this.previewModal.show = true;
                        } catch(e) { this.showToast("Błąd pobierania pliku", 'error'); }
                    },

                    async saveSmtpConfig() { try { await this.api('smtp-config', 'PUT', this.smtpConfig); this.showToast("Konfiguracja zapisana"); } catch(e) { this.showToast(e.message, 'error'); } },
                    async testSmtpConfig() { try { const res = await this.api('smtp-config/test', 'POST', this.smtpConfig); this.showToast(res.msg); } catch(e) { this.showToast(e.message, 'error'); } },
                    async addUser() { try { if(!this.forms.user.email) return; const res = await this.api('users', 'POST', this.forms.user); this.systemUsers = await this.api('users'); this.forms.user.email = ''; this.showToast(res.message); } catch(e) { this.showToast(e.message, 'error'); } },
                    async updateUserPerms(u) { try { await this.api(`users/${u.id}/permissions`, 'PUT', {can_view_fleet: u.can_view_fleet, can_edit_fleet: u.can_edit_fleet}); this.showToast("Uprawnienia zapisane"); } catch(e) { this.showToast(e.message, 'error'); } },
                    openPasswordModal(u) { this.forms.password.userId = u.id; this.forms.password.email = u.email; this.forms.password.newPassword = ''; this.modals.password = true; },
                    async submitNewPassword() { if (!this.forms.password.newPassword) return; try { await this.api(`users/${this.forms.password.userId}/password`, 'PUT', { new_password: this.forms.password.newPassword }); this.showToast("Hasło zostało zresetowane"); this.modals.password = false; } catch(e) { this.showToast(e.message, 'error'); } },
                    async deleteUser(id) { try { await this.api(`users/${id}`, 'DELETE'); this.systemUsers = this.systemUsers.filter(u => u.id !== id); this.showToast("Konto usunięte"); } catch(e) { this.showToast(e.message, 'error'); } },
                    downloadExcel() { window.open('/api/reports/export', '_blank'); }
                }
            }).mount('#app')
        </script>
    </body>
    </html>
    """

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8001))
    uvicorn.run("fleet:app", host="0.0.0.0", port=port)
