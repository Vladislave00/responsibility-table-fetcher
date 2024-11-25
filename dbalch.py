import os
from sqlalchemy import create_engine, Column, Integer, String, ForeignKey, inspect
from sqlalchemy.orm import sessionmaker, relationship, declarative_base
from sqlalchemy.exc import SQLAlchemyError
from dotenv import load_dotenv

load_dotenv()

# Define the base class for declarative models
Base = declarative_base()

# Define the Activities model
class Activities(Base):
    __tablename__ = 'activities'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255))
    subname = Column(String(255))

# Define the Responsibilities model
class Responsibilities(Base):
    __tablename__ = 'responsibilities'
    
    id = Column(Integer, primary_key=True)
    number = Column(Integer)
    letter = Column(String(255))
    indicator = Column(String(255))
    activity_id = Column(Integer, ForeignKey('activities.id'))

    activity = relationship("Activities")

# Define the Employees_for_activities model
class EmployeesForActivities(Base):
    __tablename__ = 'employees_for_activities'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255))
    post = Column(String(255))

# Define the Employee_responsibility model
class EmployeeResponsibility(Base):
    __tablename__ = 'employee_responsibility'
    
    id = Column(Integer, primary_key=True)
    employee_id = Column(Integer, ForeignKey('employees_for_activities.id'))
    responsibility_id = Column(Integer, ForeignKey('responsibilities.id'))

    employee = relationship("EmployeesForActivities")
    responsibility = relationship("Responsibilities")

# Database class encapsulating connection and query logic
class Database:
    _engine = None
    _Session = None

    @classmethod
    def get_engine(cls):
        if cls._engine is None:
            cls._engine = create_engine(
                f"postgresql://{os.getenv('DB_USER')}:{os.getenv('DB_PASSWORD')}@{os.getenv('DB_HOST')}:{os.getenv('DB_PORT')}/{os.getenv('DB_NAME')}"
            )
            print("Соединение с PostgreSQL установлено")
        return cls._engine

    @classmethod
    def get_session(cls):
        if cls._Session is None:
            cls._Session = sessionmaker(bind=cls.get_engine())
            print("Сессия создана")
        return cls._Session()
    
    @classmethod
    def list_tables(cls):
        inspector = inspect(cls.get_engine())
        tables = inspector.get_table_names()
        print("Таблицы в базе данных:", tables)

    @classmethod
    def get_data(cls):
        session = cls.get_session()  # Create a session instance
        try:
            results = session.query(
                Activities.name.label('activity_name'),
                Activities.subname,
                Responsibilities.number,
                Responsibilities.letter,
                Responsibilities.indicator,
                EmployeesForActivities.name.label('employee_name'),
                EmployeesForActivities.post.label('employee_post')
            ).outerjoin(Responsibilities, Responsibilities.activity_id == Activities.id) \
             .outerjoin(EmployeeResponsibility, EmployeeResponsibility.responsibility_id == Responsibilities.id) \
             .outerjoin(EmployeesForActivities, EmployeeResponsibility.employee_id == EmployeesForActivities.id) \
             .order_by(Responsibilities.number, Responsibilities.letter) \
             .all()
            return results
        except SQLAlchemyError as e:
            print("Ошибка при выполнении запроса:", e)
            return []
        finally:
            session.close()  # Ensure the session is closed

if __name__ == "__main__":
    db = Database()
    db.list_tables()  # List tables after connecting
    Base.metadata.create_all(db.get_engine())  # Create tables if necessary
    data = db.get_data()
    for row in data:
        print(row)
