---
title: Наборы обязательных элементов API удостоверений
description: Сведения о наборе требований API удостоверений для надстроек Office.
ms.date: 04/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 5bface00e0ffe89e7a403b251129867b334f7f69
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094381"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы обязательных элементов API удостоверений, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

|  Набор обязательных элементов  | Office 2013 или более поздней версии для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке Microsoft 365) |  Office для iPad<br>(подключено к подписке Microsoft 365)  |  Office для Mac<br>(подключено к подписке Microsoft 365)  | Office в Интернете  | SharePoint Online | OneDrive.com |Outlook.com и Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Предварительный просмотр IdentityAPI  | Недоступно | Предварительный просмотр<b>*</b> | Скоро | Предварительный просмотр<b>*</b> | Предварительный просмотр<b>* &#8224;</b> | Предварительный просмотр<b>* &#8224;</b>| Скоро | Скоро |

> **&#42;** На этапе предварительной версии API удостоверений требуется подписка на Microsoft 365. Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки. Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://insider.office.com). Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.
>
> **&#8224;** Надстройки, использующие API единого входа на этих платформах, будут работать только в том случае, если администратор клиента предоставил согласие на надстройку. Пользователь не может предоставить согласие даже в свой профиль Azure AD.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>Предварительный просмотр IdentityAPI

Подробнее об этом API можно узнать в версии, использующей обещания в [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) , или в версии, использующей функции обратного вызова по адресу [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
