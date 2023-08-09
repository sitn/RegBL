from sqlalchemy import Column, BigInteger, Integer, Float, String, Date
from sqlalchemy.ext.declarative import declarative_base


Base = declarative_base()

class RegBLApurementFeedbackHebdoCanton(Base):
    __tablename__ = 'regbl_apurement_feedback_hebdo_canton'
    __table_args__ = {'schema': 'mensuration'}
    id = Column(BigInteger, primary_key=True)
    date_version = Column(Date)
    batiments = Column(Integer)
    entrees = Column(Integer)
    liste_1 = Column(Integer)
    liste_1_pc = Column(Float)
    liste_2 = Column(Integer)
    liste_2_pc = Column(Float)
    liste_3 = Column(Integer)
    liste_3_pc = Column(Float)
    liste_4 = Column(Integer)
    liste_4_pc = Column(Float)
    liste_5 = Column(Integer)
    liste_5_pc = Column(Float)
    liste_6 = Column(Integer)
    liste_6_pc = Column(Float)
    ext_communes_validees = Column(Integer)
    ext_communes_validees_pc = Column(Float)
    batiments_manquants = Column(Integer)
    ext_batiments = Column(Integer)
    ext_batiments_gklas = Column(Integer)
    ext_batiments_gklas_pc = Column(Float)
    ext_batiments_gbaup = Column(Integer)
    ext_batiments_gbaup_pc = Column(Float)
    ext_batiments_surf30_batiments = Column(Integer)
    ext_batiments_surf30_gklas = Column(Integer)
    ext_batiments_surf30_gklas_pc = Column(Float)
    ext_batiments_surf30_gbaup = Column(Integer)
    ext_batiments_surf30_gbaup_pc = Column(Float)


class RegBLApurementFeedbackHebdoCommunes(Base):
    __tablename__ = 'regbl_apurement_feedback_hebdo_communes'
    __table_args__ = {'schema': 'mensuration'}
    id = Column(BigInteger, primary_key=True)
    no_commune_ofs = Column(Integer)
    commune = Column(String(50))
    batiments = Column(Integer)
    entrees = Column(Integer)
    batiments_manquants = Column(Integer)
    liste_1 = Column(Integer)
    liste_1_pc = Column(Float)
    liste_2 = Column(Integer)
    liste_2_pc = Column(Float)
    liste_3 = Column(Integer)
    liste_3_pc = Column(Float)
    liste_4 = Column(Integer)
    liste_4_pc = Column(Float)
    liste_5 = Column(Integer)
    liste_5_pc = Column(Float)
    liste_6 = Column(Integer)
    liste_6_pc = Column(Float)
    total_listes_pc = Column(Float)
    ext_batiments = Column(Integer)
    ext_batiments_gklas = Column(Integer)
    ext_batiments_gklas_pc = Column(Float)
    ext_batiments_gbaup = Column(Integer)
    ext_batiments_gbaup_pc = Column(Float)
    ext_batiments_surf30_batiments = Column(Integer)
    ext_batiments_surf30_gklas = Column(Integer)
    ext_batiments_surf30_gklas_pc = Column(Float)
    ext_batiments_surf30_gbaup = Column(Integer)
    ext_batiments_surf30_gbaup_pc = Column(Float)
    date_version = Column(Date)



