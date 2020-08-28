---
title: Наборы обязательных элементов API удостоверений
description: Сведения о наборе требований API удостоверений для надстроек Office.
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c2c6ea449cef08248a9ba79051b7c0c5f9baa600
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293543"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке. Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы обязательных элементов API удостоверений, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий приложений Office.

|  Набор обязательных элементов  | Office 2013 или более поздней версии для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1,3  | Недоступно | 2008 (сборка 13127,20000) или более поздняя | Скоро | 16,40 или более поздняя версия | Август, 2020 * |

> \* Изначально набор требований поддерживается в Office в Интернете только для документов, открытых из SharePoint Online и OneDrive.com. Поддержка других документов будет поступать в Office в Интернете позже в 2020.

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
- [Указание приложений Office и требований к API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
