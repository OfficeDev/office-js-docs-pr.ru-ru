# <a name="dialog-api-requirement-sets"></a>Наборы требований API общих диалогов

Наборы требований  — это именованные группы обязательных элементов API. С помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API. требованиеДополнительные сведения см. в статье [Версии Office и наборы обязательных требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы обязательных элементов API общих диалогов, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

|  Набор обязательных требований  | Office 2013 для Windows | Office 2016 для Windows (установка с помощью MSI)   | Office 2016 для Windows (установка с помощью C2R)   |  Office для iPad  |  Office 365 для Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Сборка 15.0.4855.1000 или более поздняя | Сборка 16.0.4390.1000 или более поздняя | Версия 1602 (сборка 6741.0000) или более поздняя | 1.22 или более поздняя | 15.20 или более поздняя| Январь 2017 г. | Версия 1608 (сборка 7601.6800) или более поздняя|

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Какую версию Office Я использую?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы обязательных требований API для Office

Сведения об общих наборах обязательных требований API см. в статье [Стандартные наборы обязательных элементов API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>Dialog API 1.1 

Dialog API 1.1 является первой версией API-интерфейса. Дополнительные сведения об этом API см. в справочной теме [API общих диалогов](/javascript/api/office/office.ui).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
