---
title: Наборы обязательных элементов API удостоверений
description: ''
ms.date: 05/08/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c531fd6f7c77a02402bdd70b4f2cb091787a6a52
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952189"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы обязательных элементов API удостоверений, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

|  Набор обязательных элементов  | Office 2013 или более поздней версии в Windows<br>(одноразовая покупка) | Office в Windows<br>(подключено к Office 365) |  Office для iPad<br>(подключено к Office 365)  |  Office для Mac<br>(подключено к Office 365)  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com и Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | Н/Д | Предварительный просмотр<b>*</b> | Скоро | Предварительный просмотр<b>*</b> | Предварительный просмотр<b>*</b> | Предварительный просмотр<b>*</b>| Скоро | Скоро |

> **_амп_ # 42;** На этапе предварительной версии для API удостоверений требуется Office 365 (версия подписки Office). Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки. Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1

IdentityAPI 1.1 для единого входа — это первая версия API. Дополнительные сведения об этом API см. в разделе [Справочные материалы по API единого входа](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) статьи [Включение единого входа в надстройке](/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
