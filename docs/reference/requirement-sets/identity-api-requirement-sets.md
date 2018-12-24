---
title: Наборы обязательных элементов API удостоверений
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 43a220cfada5883f292edd13cc753dc6c70e3504
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433924"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы обязательных элементов API удостоверений, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

|  Набор обязательных элементов  | Office 2013 для Windows | Office 365 для Windows   |  Office 365 для iPad  |  Office 365 для Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com и Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | Н/Д | Предварительная версия **&#42;** | Скоро | Предварительная версия **&#42;**| Предварительная версия | Предварительная версия| Скоро | Скоро |

> **&#42;** На этапе предварительной версии API удостоверений поддерживается в Windows 2016 и Mac только для пользователей, участвующих в программе предварительной оценки с ранним доступом. Чтобы присоединиться к программе предварительной оценки, см. сайт [программы предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Чтобы переключиться на ранний доступ, см. статью [Предварительная оценка — ранний доступ](https://answers.microsoft.com/ru-RU/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview).

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы обязательных элементов API для Office

Сведения о наборах обязательных элементов общего API для Office см. в [этой статье](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

IdentityAPI 1.1 для единого входа — это первая версия API. Дополнительные сведения об этом API см. в разделе [Справочные материалы по API единого входа](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) статьи [Включение единого входа в надстройке](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
