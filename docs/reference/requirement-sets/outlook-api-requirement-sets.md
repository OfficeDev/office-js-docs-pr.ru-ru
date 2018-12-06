# <a name="outlook-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Outlook

Для надстроек Outlook требуются определенные версии API, которые указываются в элементе [Requirements](/office/dev/add-ins/reference/manifest/requirements) их [манифестов](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Надстройки Outlook всегда включают элемент [Set](/office/dev/add-ins/reference/manifest/set), где для атрибута `Name` задано значение `Mailbox`, а в атрибуте `MinVersion` указан минимальный набор обязательных элементов API, поддерживающий сценарии надстройки.

Например, в следующем фрагменте манифеста указан минимальный набор обязательных элементов 1.1:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Все API Outlook относятся к [набору обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements) `Mailbox`. Существуют разные версии набора обязательных элементов `Mailbox`. Каждый новый набор API, который мы выпускаем, относится к более высокой версии набора. Не все клиенты Outlook поддерживают самый новый набор API, но если для клиента Outlook объявлена поддержка определенного набора обязательных элементов, то он поддерживает все API из этого набора.

Задайте версию минимального набора обязательных элементов в манифесте, чтобы указать клиент Outlook, в котором появится надстройка. Если клиент не поддерживает минимальный набор обязательных элементов, он не загружает надстройку. Например, если указана версия набора обязательных элементов 1.3, надстройка не отобразится в каком-либо клиенте Outlook, который не поддерживает версии 1.3. и ниже

## <a name="using-apis-from-later-requirement-sets"></a>Использование API из наборов обязательных элементов более поздних версий

Установка набора обязательных элементов не ограничивает доступные API, которые может использовать надстройка. Например, если для надстройки указан набор обязательных элементов 1.1, но она выполняется в клиенте Outlook, который поддерживает версию 1.3, надстройка может использовать API из набора обязательных элементов 1.3.

Чтобы использовать более новые API, разработчики могут просто проверить их наличие с помощью стандартных методов JavaScript.

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Такие проверки не нужно выполнять для API-интерфейсов, присутствующих в версии набора обязательных элементов, указанной в манифесте.

## <a name="choosing-a-minimum-requirement-set"></a>Выбор минимального набора обязательных элементов

Разработчикам следует использовать набор обязательных элементов самой ранней версии, содержащий набор критически важных API для сценария их работы, без которого надстройка не будет работать.

## <a name="clients"></a>Клиенты

Указанные ниже клиенты поддерживают надстройки Outlook.

| Клиент | Поддерживаемые наборы обязательных элементов API |
| --- | --- |
| Outlook 2019 для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 для Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 ("нажми и работай") для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 (MSI) для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2016 для Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2013 для Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook для iPhone | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook для Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook в Интернете (Office 365 и Outlook.com) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook Web App (локальная версия Exchange 2013) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |
| Outlook Web App (локальная версия Exchange 2016) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| Outlook Web App (локальная версия Exchange 2019) | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |

> [!NOTE]
> Поддержка версии 1.3 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3114349) от 8 декабря 2015 г.](https://support.microsoft.com/kb/3114349) Поддержка версии 1.4 в Outlook 2013 добавлена в рамках [обновления для Outlook 2013 (KB3118280) от 13 сентября 2016 г.](https://support.microsoft.com/help/3118280) Поддержка версии 1.4 в Outlook 2016 (MSI) добавлена в рамках [обновления для Office 2016 (KB4022223) от 3 июля 2018 г.](https://support.microsoft.com/help/4022223).
