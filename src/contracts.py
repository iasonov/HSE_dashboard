"""Bitrix admissions data contracts used by the dashboard pipeline."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class BitrixEntity:
    """Describe one Bitrix admissions entity table."""

    name: str
    id_field: str
    required_fields: tuple[str, ...]


BITRIX_DEALS = BitrixEntity(
    name="deals",
    id_field="idaispk",
    required_fields=(
        "idaispk",
        "idcontact",
        "idop",
        "date_registrationaispk",
        "date_dogovora",
        "prikaz_zachislenya", # TODO reg_nomer, date_zayvlenya_soglasie, dogovor_oplachen, tip_oplaty, prioritet_plat_mest, skidka_rezultatvi, campus, tekuroven_obrazovanya, vid_mesta
    ),
)
BITRIX_CONTACTS = BitrixEntity(
    name="contacts",
    id_field="idaispk",
    required_fields=("idaispk", "pol", "birthdate"), # TODO idgrazhdanstvo, idstrana_prozhivanya, inostranec

)
BITRIX_EDUCATIONAL_PROGRAMS = BitrixEntity(
    name="educational_programs",
    id_field="idaispk",
    required_fields=("idaispk", "name", "uroven_obrazovanya", "campus"), # TODO tip_op, facultet

)
BITRIX_CONTRACTS = BitrixEntity(
    name="contracts",
    id_field="idaispk",
    required_fields=("idaispk", "iddeal", "data_oplaty"), # TODO istochnik, datetimecreate, idregion_prozhivanya
)
BITRIX_EXAMS = BitrixEntity(
    name="exams",
    id_field="idaispk",
    required_fields=("idaispk", "idcontact", "iddeal", "ball", "date_testirovanya", "aktive"),
)
BITRIX_PORTFOLIOS = BitrixEntity(
    name="portfolios",
    id_field="idaispk",
    required_fields=("idaispk", "idcontact", "iddeal", "idtovar", "status_elementa_portfolio"),
)

BITRIX_ADMISSIONS_ENTITIES: tuple[BitrixEntity, ...] = (
    BITRIX_DEALS,
    BITRIX_CONTACTS,
    BITRIX_EDUCATIONAL_PROGRAMS,
    BITRIX_CONTRACTS,
    BITRIX_EXAMS,
    BITRIX_PORTFOLIOS,
)

