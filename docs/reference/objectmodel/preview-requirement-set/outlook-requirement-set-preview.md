---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 06/03/2019
localization_priority: Priority
ms.openlocfilehash: 6d7dceba80929c2abace9daec859302c43913f98
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706353"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки. Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность. Также может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider).

Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

### <a name="attachments"></a>Вложения

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Добавлен новый объект, представляющий содержимое вложения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

Добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

Добавлен новый метод, позволяющий получить содержимое определенного вложения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

Добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

Добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

Добавлено событие `AttachmentsChanged` в объект `Item`.

**Доступно в** Outlook для Windows (подключенный к Office 365)

---

### <a name="block-on-send"></a>Блокировка при отправке

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

Добавлен новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением `allowEvent`. Это значение используется для отмены выполнения события.

**Доступно в** Outlook в Интернете (классическая версия)

---

### <a name="categories"></a>Категории

В Outlook пользователь может группировать сообщения и встречи, используя категории для выделения их цветом. Пользователь определяет категории в главном списке своего почтового ящика. Затем он может применить одну или несколько категорий к элементу.

> [!NOTE]
> Эта функция не поддерживается в Outlook для iOS или Outlook для Android.

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[Categories](/javascript/api/outlook/office.categories)

Добавлен новый объект, представляющий категории элемента.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[CategoryDetails](/javascript/api/outlook/office.categorydetails)

Добавлен новый объект, представляющий сведения о категории (ее имя и соответствующий цвет).

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[MasterCategories](/javascript/api/outlook/office.mastercategories)

Добавлен новый объект, представляющий главный список категорий для почтового ящика.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)

Добавлено новое свойство, представляющее главный список категорий для почтового ящика.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[Office.context.mailbox.item.categories](/javascript/api/outlook/office.item#categories)

Добавлено новое свойство, представляющее набор категорий для элемента.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor)

Добавлено новое перечисление, указывающее цвета, доступные для сопоставления с категориями.

**Доступно в** Outlook для Windows (подключенный к Office 365)

---

### <a name="delegate-access"></a>Делегированный доступ

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

Добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback)

Добавлен новый метод, который получает идентификатор сохраненного элемента встречи или сообщения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

Добавлен новый метод, позволяющий получить объект, который представляет свойства sharedProperties элемента встречи или сообщения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

Добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[Элемент манифеста SupportsSharedFolders](../../manifest/supportssharedfolders.md)

К элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент. Он определяет, доступна ли надстройка в сценариях делегирования.

**Доступно в** Outlook для Windows (подключенный к Office 365)

---

### <a name="enhanced-location"></a>Расширенные функции расположения

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Добавлен новый объект, представляющий набор расположений для встречи.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

Добавлен новый объект, представляющий расположение. Только для чтения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

Добавлен новый объект, представляющий идентификатор расположения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

Добавлено новое свойство, представляющее набор расположений для встречи.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

Добавлено новое перечисление, которое определяет тип расположения встречи.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

Добавлено событие `EnhancedLocationsChanged` в объект `Item`.

**Доступно в** Outlook для Windows (подключенный к Office 365)

---

### <a name="integration-with-actionable-messages"></a>Взаимодействие с интерактивными сообщениями

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Доступно в** Outlook для Windows (подключенный к Office 365), Outlook в Интернете (классическая версия)

---

### <a name="internet-headers"></a>Заголовки Интернета

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

Добавлен новый объект, представляющий заголовки Интернета в элементе сообщения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

Добавлено новое свойство, представляющее заголовки Интернета в элементе сообщения.

**Доступно в** Outlook для Windows (подключенный к Office 365)

---

### <a name="office-theme"></a>Тема Office

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[Office.context.mailbox.officeTheme](/javascript/api/office/office.officetheme)

Добавлена возможность получения темы Office.

**Доступно в** Outlook для Windows (подключенный к Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Добавлено событие `OfficeThemeChanged` в объект `Mailbox`.

**Доступно в** Outlook для Windows (подключенный к Office 365)

---

### <a name="sso"></a>Единый вход

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.

**Доступно в** Outlook для Windows (подключенный к Office 365), Outlook для Mac (подключенный к Office 365), Outlook в Интернете (Outlook.com и подключенный к Office 365), Outlook в Интернете (классическая версия)

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/outlook/add-ins/quick-start)
