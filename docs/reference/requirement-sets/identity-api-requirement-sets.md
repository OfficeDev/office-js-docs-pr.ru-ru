---
title: Наборы обязательных элементов API удостоверений
description: API удостоверений заданная информация для Office надстройки.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: bff7d75d538922f6d5d5d05a01306a4ba2ec836c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744930"
---
# <a name="identity-api-requirement-sets"></a>Наборы обязательных элементов API удостоверений

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API удостоверений, Office клиентских приложений, поддерживаюющих этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2021 или более поздней Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(обе подписки<br> и разовая покупка Office Mac 2019 и более поздних периодов)   | Office в Интернете  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | Сборка 16.0.14326.20454 или более поздней | Версия 2008 (сборка 13127.20000) или более поздней версии | Не поддерживается | 16.40 или более поздняя | Microsoft Office SharePoint Online и OneDrive\* |

\*В настоящее время набор требований поддерживается в Office в Интернете только для документов, которые открываются из Microsoft Office SharePoint Online и OneDrive.

## <a name="outlook-and-identity-api-requirement-sets"></a>Outlook и наборы API удостоверений

[!INCLUDE [How to use the Identity 1.3 requirement set in Outlook add-ins](../../includes/outlook-identity-13-note.md)]

> [!NOTE]
> В Outlook с помощью активации на основе событий интерфейс [OfficeRuntime.Auth](/javascript/api/office-runtime/officeruntime.auth) поддерживается в Office версии Windows версии 2108 (сборка 14326.20258) или более поздней версии. В [Office. Интерфейс Auth](/javascript/api/office/office.auth) поддерживается в версии 2109 (сборка 14425.10000) или более поздней версии. Дополнительные сведения в соответствии с вашей версией см. на странице история обновления Office [2021](/officeupdates/update-history-office-2021) или [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) и поиске клиентской версии Office и канала [обновления](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
