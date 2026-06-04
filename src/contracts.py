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
        # "idaispk",
        # "idcontact",
        # "idop",
        # "date_registrationaispk",
        # "date_dogovora",
        # "prikaz_zachislenya", # TODO reg_nomer, date_zayvlenya_soglasie, dogovor_oplachen, tip_oplaty, prioritet_plat_mest,
        #  # skidka_rezultatvi, campus, tekuroven_obrazovanya, vid_mesta
        #"title" - фио
        #"comments" - комментарии
        "createdTime",
        "updatedTime",
        "begindate",
        "closedate",
        "createdBy",
        "updatedBy",
        "contactId",
        "assignedById", # id менеджера?
        "categoryId",
        "stageId", # стадия воронки
        "stageSemanticId",
        "productId",
        "isNew",
        "closed",
        "sourceId",
        "lastCommunicationTime", # TODO check последнее ли это время связи
        "ufDealUrlOp",
        "ufDealYearAdmission", # год подачи?
        "ufDealRanneePriglashenie", # РП?
        "ufDealUtmSource", # +Medium, Campaign, Content, Term
        "ufDealFormname",
        "ufDEAL_OP_FORCHAT2DESK",
        "ufDealDogovorOplachen",
        "ufDealContractdate", # !!! дата договора
        "ufDealPrikazOZachislenii", # есть ли внутри дата приказа?
        "ufDealResultatRassmotrenya", # ??
        "ufDealRegNomer",
        "ufDealUinAsav",
        "ufDealUinAispk",
        "ufDealPrioritetKommercMesto", # ufDealPrioritetBudjetMesto, ufDealPrioritetCelevoeMesto
        "ufDealFinancing",
        "ufDealDataPodachiSoglasia", # можно использовать как дату оплаты?
        "ufDealCampus",
        "ufDealRekomKZachislen",
        "ufDealPortfolio",
        "ufDealZachislen",
        "ufDealSendingStatus", # список из id рассылок?
        "utmSource", # +Medium,...
        "searchContent", # '9118 Савченко Мария 0.00 Российский рубль Рылова Татьяна Рылова Беляева Мария Викторовна 79854584331 89854584331 9854584331 854584331 54584331 4584331 584331 84331 4331 331 ziorylnrin_5 rqh ufr eh Продажа Первый контакт 06.09.2025 13.09.2025 [c]\n22.04.2026 оставила повторную заявку, в этом году заканчивает бак по Юриспруденции планирует поступать. Спрашивала насколько сложные ВИ, каков процент их сдачи, диплом получает в начале июня, будет готовиться.\n08.09.2025 не взяла трубку\n18.08 ндз\n08.09 Арина Щ. входящий. в этом учебном году заканчивает бакалавриат вышки, хочет поступить на бакалавриат фкн, уточнила, может ли это сделать по вступительным - да. интересует кнад, и программирование, хочет второе базовое образование, а потом уже в магистратуру на искинт. фанат вышки и хочет долго у нас учиться)) отправляю инфо, будет ждать связи\n[/c] bayvar.ufr.eh 22.04.2026 18:30:26 Бакалавриат КНАД. Компьютерные науки и анализ данных / Москва / 010302 Прикладная математика и информатика / факультет компьютерных наук / Бакалавриат 2026 не выбрано Успешный звонок Недозвон 57368842 502810 uggcf://fghqlbayvar.ufr.eh/ ahyy Оставьте заявку на консультацию ДОО. Отправлено письмо недозвон ДОО. Отправлено приветственное письмо без программы ДОО. Отправлено P2q приветственное письмо ДОО. Отправлено P2q Недозвон 1 ДОО. Отправлено письмо Успешный звонок 408854 ИТ'
        # Не найдено:
        # 'Дата проведения оплаты': 'PAYMENT_PAID',
        # 'Товар': 'PRODUCT_ROW_PRODUCT_ID',
        # 'Дата регистрации': 'UF_DEAL_DATA_REGISTRACII',
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

