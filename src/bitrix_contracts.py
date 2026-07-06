'''Bitrix admissions data contracts used by the dashboard pipeline.'''

from __future__ import annotations

from dataclasses import dataclass

from collections.abc import Mapping
from typing import Any

@dataclass(frozen=True, slots=True)
class BitrixEntity:
    '''Describe one Bitrix admissions entity table.'''

    name: str
    id_field: str # TODO check - is it needed?
    select_name: str                    # 'select' or 'SELECT' # TODO test with different case/CASE 
    filter_name: str                    # 'filter' or 'FILTER' # TODO test with different case/CASE
    select: tuple[str, ...]             # what information to get
    extra_filter: Mapping[str, Any]     # extra filters for select
    request_id: int                     # number id of table
    request_rest: str                   # name of request method
    request_base_params: dict[str, Any] # base request parameter except select and filter


# Поступление 360
BITRIX_CRM_DEALS = BitrixEntity(
    name='crm_deals',
    id_field='id',
    request_id=4, # воронка 'Поступление 360'
    select_name='select',
    filter_name='filter',
    select=(
        'id', 'createdTime', 'updatedTime', 'createdBy', 'updatedBy',
       'assignedById', 'opened', 'leadId', 'companyId', 'contactId', 'title',
       'categoryId', 'stageId', 'stageSemanticId', 'isNew', 'isRecurring',
       'isReturnCustomer', 'isRepeatedApproach', 'closed', 'typeId',
       'opportunity', 'isManualOpportunity', 'currencyId', 'comments',
       'begindate', 'closedate', 'sourceId', 'sourceDescription',
       'originatorId', 'originId', 'searchContent', 'movedBy', 'movedTime',
       'lastActivityBy', 'lastActivityTime', 'lastCommunicationTime',
       'ufCrm_1757155225', 'ufDealUrovenObrazovanya', 'ufDealEducationProgram',
       'ufDealCityOfResidence', 'ufDealNotFree', 'ufDealYearAdmission',
       'ufDealEducationFor', 'ufDealDopProgram', 'ufDealFormatProgram',
       'ufDealUspechnyCall', 'ufDealNedozvon', 'ufDEAL_NEDOZVON2',
       'ufDEAL_NEDOZVON3', 'ufDealNaborNaGod', 'ufDealUdobnyVidSvyazi',
       'ufDealRegisteredAsav', 'ufDealRoistat', 'ufDealUnsubscribe',
       'ufDealFormaObuchebya', 'ufDealRanneePriglashenie', 'ufDealIdold',
       'ufDealCommentsDubl', 'ufDealCommentsOtkaz', 'ufDealUrlOp',
       'ufDealZhelaemyUrovenObrazovanya', 'ufDealUtmSource', 'ufDealUtmMedium',
       'ufDealUtmCampaign', 'ufDealUtmContent', 'ufDealUtmTerm',
       'ufDealFormname', 'ufDEAL_OP_FORCHAT2DESK', 'ufDealVopros',
       'ufDealSendingStatus', 'ufDealPrichinaProvala', 'ufDealDatajson',
       'ufCrm_688A1B246A92A', 'ufDealNecelevoyDoo', 'ufCrm_1757147711922',
       'ufCrmInUmnico', 'ufCrmUmnicoLeadId', 'ufCrm_1761655071',
       'ufCrm_1761923285', 'ufDealPortfolio', 'ufDealZachislen',
       'ufCrm_1763561124', 'ufCrm_1766664747', 'ufDealPortfolioTech',
       'ufCrm_1770885545', 'utmSource', 'utmMedium', 'utmCampaign',
       'utmContent', 'utmTerm', 'observers', 'contactIds', 'entityTypeId',
        # 'id',
        # # 'idaispk',
        # # 'idcontact',
        # # 'idop',
        # # 'date_registrationaispk',
        # # 'date_dogovora',
        # # 'prikaz_zachislenya', # TODO reg_nomer, date_zayvlenya_soglasie, dogovor_oplachen, tip_oplaty, prioritet_plat_mest,
        # #  # skidka_rezultatvi, campus, tekuroven_obrazovanya, vid_mesta
        # #'title' - фио
        # #'comments' - комментарии
        # 'createdTime',
        # 'updatedTime',
        # 'begindate',
        # 'closedate',
        # 'createdBy',
        # 'updatedBy',
        # 'contactId',
        # 'assignedById', # id менеджера?
        # 'categoryId',
        # 'stageId', # стадия воронки
        # 'stageSemanticId',
        # 'productId',
        # 'isNew',
        # 'closed',
        # 'sourceId',
        # 'lastCommunicationTime', # TODO check последнее ли это время связи
        # 'ufDealEducationProgram', # TODO check
        # 'ufDealUrlOp',
        # 'ufDealYearAdmission', # год подачи?
        # 'ufDealRanneePriglashenie', # РП?
        # 'ufDealUtmSource', # +Medium, Campaign, Content, Term
        # 'ufDealFormname',
        # 'ufDEAL_OP_FORCHAT2DESK',
        # 'ufDealDogovorOplachen',
        # 'ufDealContractdate', # !!! дата договора
        # 'ufDealPrikazOZachislenii', # есть ли внутри дата приказа?
        # 'ufDealResultatRassmotrenya', # ??
        # 'ufDealRegNomer',
        # 'ufDealUinAsav',
        # 'ufDealUinAispk',
        # 'ufDealPrioritetKommercMesto', # ufDealPrioritetBudjetMesto, ufDealPrioritetCelevoeMesto
        # 'ufDealFinancing',
        # 'ufDealDataPodachiSoglasia', # можно использовать как дату оплаты?
        # 'ufDealCampus',
        # 'ufDealRekomKZachislen',
        # 'ufDealPortfolio',
        # 'ufDealZachislen',
        # 'ufDealSendingStatus', # список из id рассылок?
        # 'utmSource', # +Medium,...
        # 'searchContent', # '9118 Савченко Мария 0.00 Российский рубль Рылова Татьяна Рылова Беляева Мария Викторовна 79854584331 89854584331 9854584331 854584331 54584331 4584331 584331 84331 4331 331 ziorylnrin_5 rqh ufr eh Продажа Первый контакт 06.09.2025 13.09.2025 [c]\n22.04.2026 оставила повторную заявку, в этом году заканчивает бак по Юриспруденции планирует поступать. Спрашивала насколько сложные ВИ, каков процент их сдачи, диплом получает в начале июня, будет готовиться.\n08.09.2025 не взяла трубку\n18.08 ндз\n08.09 Арина Щ. входящий. в этом учебном году заканчивает бакалавриат вышки, хочет поступить на бакалавриат фкн, уточнила, может ли это сделать по вступительным - да. интересует кнад, и программирование, хочет второе базовое образование, а потом уже в магистратуру на искинт. фанат вышки и хочет долго у нас учиться)) отправляю инфо, будет ждать связи\n[/c] bayvar.ufr.eh 22.04.2026 18:30:26 Бакалавриат КНАД. Компьютерные науки и анализ данных / Москва / 010302 Прикладная математика и информатика / факультет компьютерных наук / Бакалавриат 2026 не выбрано Успешный звонок Недозвон 57368842 502810 uggcf://fghqlbayvar.ufr.eh/ ahyy Оставьте заявку на консультацию ДОО. Отправлено письмо недозвон ДОО. Отправлено приветственное письмо без программы ДОО. Отправлено P2q приветственное письмо ДОО. Отправлено P2q Недозвон 1 ДОО. Отправлено письмо Успешный звонок 408854 ИТ'
        # # Не найдено:
        # # 'Дата проведения оплаты': 'PAYMENT_PAID',
        # # 'Товар': 'PRODUCT_ROW_PRODUCT_ID',
        # # 'Дата регистрации': 'UF_DEAL_DATA_REGISTRACII',
        # # 'Программа': 'UF_DEAL_EDUCATION_PROGRAM'
    ),
    extra_filter={'CATEGORY_ID' : 4, '>=begindate' : '2025-10-01'},
  
    request_rest='crm.item.list',
    request_base_params={
        'entityTypeId': 2,  # 2 - сделки
        # 'select': ['*'], # TODO list(select), now for the case of problem
        # 'filter': dict(extra_filter),
        'order': {'id': 'ASC'},
    },
)

# Портал ВШЭ
BITRIX_PORTAL_DEALS = BitrixEntity(
    name='portal_deals',
    id_field='id',
    select_name='select',
    filter_name='filter',
    select=('id', 'createdTime', 'updatedTime', 'createdBy', 'updatedBy',
       'assignedById', 'opened', 'companyId', 'contactId', 'title',
       'categoryId', 'stageId', 'stageSemanticId', 'isNew', 'isRecurring',
       'isReturnCustomer', 'closed', 'typeId', 'currencyId', 'begindate',
       'closedate', 'sourceId', 'searchContent', 'movedTime', 'lastActivityBy',
       'lastActivityTime', 'lastCommunicationTime', 'ufCrm_1757155225',
       'ufDealEducationProgram', 'ufDealDopProgram', 'ufDealNaborNaGod',
       'ufDealRoistat', 'ufDealIdold', 'ufDealUrlOp', 'ufDealVopros',
       'ufDealDatajson', 'ufCrm_1757147711922', 'ufDealDopNabor',
       'ufCrm_1761923285', 'ufDealPortfolio', 'ufCrm_1763561124',
       'contactIds',), # BITRIX_CRM_DEALS.select,
    extra_filter={'CATEGORY_ID' : 2, '>=begindate' : '2025-10-01'},
    
    request_id=2, # 2 - Воронка 'Портал ВШЭ'
    request_rest='crm.item.list',
    request_base_params={
        'entityTypeId': 2,  # 2 - сделки
        # 'select': ['*'], # TODO list(select), now for the case of problem
        # 'filter': dict(extra_filter),
        'order': {'id': 'ASC'},
    },
)

# АСАВ и АИС ПК - маг/бак
BITRIX_APPLICATIONS = BitrixEntity(
    name='applications',
    id_field='id',
    select_name='select',
    filter_name='filter',
    select=('id', 'createdTime', 'updatedTime', 'createdBy', 'updatedBy',
       'assignedById', 'opened', 'companyId', 'contactId', 'title',
       'categoryId', 'stageId', 'stageSemanticId', 'isNew', 'isRecurring',
       'isReturnCustomer', 'closed', 'typeId', 'currencyId', 'comments',
       'begindate', 'closedate', 'sourceId', 'searchContent', 'movedBy',
       'movedTime', 'lastActivityBy', 'lastActivityTime',
       'lastCommunicationTime', 'ufCrm_1757155225', 'ufDealEducationProgram',
       'ufDealDopProgram', 'ufDealNaborNaGod', 'ufDealRegisteredAsav',
       'ufDealRoistat', 'ufDealFormaObuchebya', 'ufDealIdold', 'ufDealUrlOp',
       'ufDealZhelaemyUrovenObrazovanya', 'ufDEAL_OP_FORCHAT2DESK',
       'ufDealNomerDogovora', 'ufDealDogovorOplachen', 'ufDealContractdate',
       'ufDealResultatRassmotrenya', 'ufDealDeystvZayavka',
       'ufDealPrikazOZachislenii', 'ufDealVseOcenkiUdovletvorit',
       'ufDealSummaBallov', 'ufDealRegNomer', 'ufDealStatusKontakta',
       'ufDealSendingStatus', 'ufDealUinAsav', 'ufDealUinAispk',
       'ufDealPrioritetBudjetMesto', 'ufDealPrioritetKommercMesto',
       'ufDealPrioritetCelevoeMesto', 'ufDealDataRegistracii', 'ufDealCampus',
       'ufDealFinancing', # бюджетные места
       'ufDealDataPodachiSoglasia',
       'ufDealPredMestoObuchenia', 'ufDealDogovorPodpisan', 'ufDealChild',
       'ufDealRekomKZachislen', 'ufDealPrichinaProvala', 'ufCrm_1757147711922',
       'ufDealDopNabor', 'ufDealOsnovanieZachislenVibitiya',
       'ufCrm_1761655071', 'ufCrm_1761923285', 'ufDealPortfolio',
       'ufDealZachislen', 'ufCrm_1763561124', 'utmSource', 'utmMedium',
       'utmCampaign', 'utmContent', 'utmTerm', 'contactIds', 'entityTypeId',), # BITRIX_CRM_DEALS.select,
    extra_filter={'CATEGORY_ID' : 1, '>=begindate' : '2026-06-20'},
    
    request_id=1, # 1- Воронка 'ОНЛАЙН МАГ/БАК'
    request_rest='crm.item.list',
    request_base_params={
        'entityTypeId': 2,  # 2 - сделки
        # 'select': ['*'], # TODO list(select), now for the case of problem
        # 'filter': dict(extra_filter),
        'order': {'id': 'ASC'},
    },
)

BITRIX_CONTACTS = BitrixEntity(
    name='contacts',
    id_field='id',
    select_name='select',
    filter_name='filter',
    request_id=3, #TODO check
    select=('id', 'createdTime', 'updatedTime', 'createdBy', 'updatedBy',
       'assignedById', 'opened', 'companyId', 'sourceId', 'sourceDescription',
       'name', 'lastName', 'secondName', 'photo', 'post', 'address',
       'comments', 'leadId', 'export', 'typeId', 'webformId', 'originatorId',
       'originId', 'birthdate', 'hasPhone', 'hasEmail', 'hasImol',
       'searchContent', 'lastActivityBy', 'lastActivityTime', 'emailHome',
       'emailWork', 'phoneMobile', 'phoneWork', 'imol', 'email', 'phone',
       'lastCommunicationTime', 'ufContactPol', 'ufContactPriznakRanprigp',
       'ufContactSoglasieObrabotkaPdn', 'ufContactSoglasieReklama',
       'ufContactIdold', 'ufContactUinAispk', 'ufContactRanneePriglashenie',
       'ufContactUinasav', 'ufCrm_1757267473', 'ufCrmGeoUniUnsubscribe',
       'ufContactStateZayavkiRanprig', 'ufContactPrimenenoDrugoiKonkurs',
       'utmSource', 'utmMedium', 'utmCampaign', 'utmContent', 'utmTerm',
       'observers', 'companyIds', 'entityTypeId', 'fm',
        # 'id',
        # 'createdTime',
        # 'updatedTime',
        # 'createdBy',
        # 'updatedBy',
        # 'assignedById', # id менеджера 
        # 'sourceId',
        # 'name',
        # 'comments',
        # 'leadId',
        # 'categoryId',
        # #'pol', 'birthdate',
        # 'utmSource', # +Medium,...
        # 'ufContactUinAispk',
        # 'ufContactRanneePriglashenie',
    ), # TODO idgrazhdanstvo, idstrana_prozhivanya, inostranec
    
    extra_filter={},
  
    request_rest='crm.item.list',
    request_base_params={
            'entityTypeId': 3,  # 3 - контакты
            # 'select': ['*'], # TODO list(select), now for the case of problem
            # 'filter': dict(extra_filter),
            'order': {'id': 'ASC'},
    },
)

# https://apidocs.bitrix24.ru/api-reference/lists/elements/lists-element-get.html 
BITRIX_EDUCATIONAL_PROGRAMS = BitrixEntity(
    name='educational_programs',
    id_field='id',
    select_name='SELECT',
    filter_name='FILTER',
    select=('ID', 'IBLOCK_ID', 'NAME', 'CREATED_BY', 'BP_PUBLISHED', 'DATE_CREATE',
       'CREATED_USER_NAME', 'TIMESTAMP_X', 'MODIFIED_BY', 'USER_NAME',
       'PROPERTY_80', 'PROPERTY_81', 'PROPERTY_82', 'PROPERTY_83',
       'PROPERTY_84', 'PROPERTY_90', 'PROPERTY_97', 'PROPERTY_103',
       'PROPERTY_104', 'PROPERTY_105', 'PROPERTY_106', 'PROPERTY_107',
       'PROPERTY_108', 'PROPERTY_109', 'PROPERTY_110', 'PROPERTY_111',
       'PROPERTY_184', 'PROPERTY_228', 'PROPERTY_625', 'PROPERTY_637',
       'PROPERTY_638',), # TODO tip_op, facultet  'uroven_obrazovanya', 'campus'
    extra_filter={},
    request_id=21, #TODO check
    request_rest='lists.element.get', 
    request_base_params={
        'IBLOCK_TYPE_ID' : 'lists',
        'IBLOCK_ID': 21, #request_id
        # 'SELECT' : ('ID', 'NAME',), #select #TODO Check
        # 'FILTER' : dict(extra_filter), 
        'ELEMENT_ORDER': {'id': 'ASC'},
    },
)
# BITRIX_CONTRACTS = BitrixEntity(
#     name='contracts',
#     id_field='idaispk',
#     required_fields=('idaispk', 'iddeal', 'data_oplaty'), # TODO istochnik, datetimecreate, idregion_prozhivanya
# )
# BITRIX_EXAMS = BitrixEntity(
#     name='exams',
#     id_field='idaispk',
#     required_fields=('idaispk', 'idcontact', 'iddeal', 'ball', 'date_testirovanya', 'aktive'),
# )
# BITRIX_PORTFOLIOS = BitrixEntity(
#     name='portfolios',
#     id_field='idaispk',
#     required_fields=('idaispk', 'idcontact', 'iddeal', 'idtovar', 'status_elementa_portfolio'),
# )

BITRIX_ADMISSIONS_ENTITIES: tuple[BitrixEntity, ...] = (
    BITRIX_CRM_DEALS,
    BITRIX_PORTAL_DEALS,
    BITRIX_APPLICATIONS,
    # BITRIX_CONTACTS, # А зачем, собственно, нужны контакты? Там только ДР да и то не везде
    BITRIX_EDUCATIONAL_PROGRAMS,
    # BITRIX_CONTRACTS,
    # BITRIX_EXAMS,
    # BITRIX_PORTFOLIOS,
)

