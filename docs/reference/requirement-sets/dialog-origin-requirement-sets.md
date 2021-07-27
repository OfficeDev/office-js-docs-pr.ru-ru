---
title: Наборы требований к диалоговом происхождению
description: Дополнительные информацию о наборах требований к диалоговом происхождению.
ms.date: 07/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 1ec5949c689021f080491a19aea4661627b2d95c
ms.sourcegitcommit: f46e4aeb9c31f674380dd804fd72957998b3a532
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/23/2021
ms.locfileid: "53536065"
---
# <a name="dialog-origin-requirement-sets"></a>Наборы требований к диалоговом происхождению

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований к диалоговому источнику, Office клиентские приложения, которые поддерживают этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2013 для Windows\*<br>(единовременная покупка) | Office 2016 для Windows\*<br>(единовременная покупка) | Office 2019 или более поздней Windows\*<br>(единовременная покупка) | Office для Windows<br>(подписка) |  Office для iPad<br>(подписка)  |  Office для Mac<br>(подписка)  | Office в Интернете  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | Сборка<br>15.0.5371.1000<br>или более поздней | Сборка<br>16.0.5200.1000<br>или более поздней | Сборка<br>Подлежит уточнению.<br>или более поздней | Подлежит уточнению. | 2.52 или более поздней | 16.52 или более поздней | Июль 2021 г. | Версия 2108<br>(Сборка 10377.1000)<br>или более поздней |

>\*Пользователи разовой системы Office, возможно, не приняли все исправления и обновления. Если это так, то DLL, Office для отчета о своей версии в пользовательском интерфейсе, может быть больше, чем указанные здесь версии, даже если обновленные DLLs, необходимые для поддержки DialogApi, не установлены на компьютере пользователя. Чтобы обеспечить установку необходимого исправления, пользователю необходимо перейти в список обновлений Office [(Office 2013](/officeupdates/msp-files-office-2013) или [Office 2016](/officeupdates/msp-files-office-2016)г.), **поискать osfclient-x-none** и установить указанный патч.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-origin-11"></a>Диалоговое начало 1.1

Диалоговое начало 1.1 — это первая версия API. Он обеспечивает поддержку меж доменного обмена сообщениями между диалогом и родительской страницей. Сведения об этих API см. в справочной [Office.ui.](/javascript/api/office/office.ui)

## <a name="see-also"></a>См. также

- [Использование Office Dialog API в надстройках Office](../../develop/dialog-api-in-office-add-ins.md)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
