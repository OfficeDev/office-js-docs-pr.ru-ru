---
title: Наборы обязательных элементов API удостоверений
description: Требования К API удостоверений устанавливают сведения для надстройок Office.
ms.date: 01/26/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c662e7a5306692fd75de51acc7cadfd1df3e7406
ms.sourcegitcommit: 85b4839be743059bf155ff44e49d64968444d80a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2021
ms.locfileid: "51471726"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API удостоверений, клиентские приложения Office, поддерживают этот набор требований, а также номера сборки или версии для приложения Office.

|  Набор обязательных элементов  | Office 2013 или более поздней версии для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | Н/Д | 2008 (сборка 13127.20000) или более поздней | Скоро | 16.40 или более поздняя | Microsoft SharePoint Online и OneDrive\* |

\* В настоящее время набор требований поддерживается в Office в Интернете только для документов, открытых в Microsoft SharePoint Online и OneDrive.

> [!NOTE]
> Outlook. Чтобы в коде надстройки был установлен API удостоверений 1.3, проверьте, поддерживается ли он путем `isSetSupported('IdentityAPI', '1.3')` вызова. Объявление его в манифесте надстройки Outlook не поддерживается. Также можно определить, поддерживается ли API, проверив, не `undefined` ли он. Подробнее см. в статье [Использование API из наборов требования более поздних версий](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>Предварительный просмотр IdentityAPI

Подробные сведения об этом API см. в версии, которая использует promises на [getAccessToken,](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) или версии, которая использует вызовы в [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
