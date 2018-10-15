# <a name="identity-api-requirement-sets"></a>Наборы требований API Identity

Наборы требований  — это именованные группы требований  API. С помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы требований ](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Запуск надстроек Office в разных версиях Office. В следующей таблице перечислены наборы требований API Identity, ведущие приложения Office, которые поддерживают наборы требований, сборку или номера версии ведущего приложения Office.

|  Набор требований  | Office 2013 для Windows | Office 365 для Windows   |  Office для iPad  |  Office 365 для Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com и Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | Н/Д | Предварительная версия ***** | Ожидается в скором времени | Предварительная версия *****| Предварительная версия | Предварительная версия| Ожидается в скором времени | Ожидается в скором времени |

> ***** На этапе предварительной  версии API-интерфейс Identity поддерживается в Windows 2016 и Mac только для пользователей программы предварительной оценки Office с использованием опции Fast. Чтобы присоединиться к программе предварительной оценки Office, обратитесь к [Стать участником программы предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Чтобы переключиться на отслеживание Fast, см. [Fast участника программы предварительной оценки Office](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Какую версию Office Я использую?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы требований  API для Office

Сведения об общих наборах требований  API см. в статье [Общие наборы требований  API для Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

IdentityAPI 1.1 для единого входа — это первая версия API. Для получения дополнительных сведений об этом API обратитесь к разделу [Справка по API службы единого входа](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) для [Включения службы единого входа в надстройку](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
