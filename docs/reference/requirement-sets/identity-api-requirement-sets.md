---
title: Наборы обязательных элементов API удостоверений
description: API удостоверений заданная информация для Office надстройки.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: e3af8767666d3015894c0b7bcdecd758b1a1547c
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/19/2021
ms.locfileid: "59450802"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API удостоверений, Office клиентских приложений, поддерживаюющих этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2021 или более поздней Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | 2008 (сборка 13127.20000) или более поздней | 2008 (сборка 13127.20000) или более поздней | Не поддерживается | 16.40 или более поздняя | Microsoft Office SharePoint Online и OneDrive\* |

\*В настоящее время набор требований поддерживается в Office в Интернете только для документов, которые открываются из Microsoft Office SharePoint Online и OneDrive.

> [!NOTE]
> Outlook. Чтобы потребовать API удостоверения 1.3 в коде надстройки, проверьте, поддерживается ли он путем `isSetSupported('IdentityAPI', '1.3')` вызова. Объявление его в манифесте Outlook надстройки не поддерживается. Также можно определить, поддерживается ли API, проверив, не `undefined` ли он. Подробнее см. в статье [Использование API из наборов требования более поздних версий](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="identityapi-preview"></a>Предварительный просмотр IdentityAPI

Подробные сведения об этом API см. в версии, которая использует promises на [getAccessToken,](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) или версии, которая использует вызовы в [getAccessTokenAsync](/javascript/api/office/office.auth#getAccessTokenAsync_options__callback_).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
