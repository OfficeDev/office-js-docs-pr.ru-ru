---
title: Наборы обязательных элементов API диалоговых окон
description: Дополнительные сведения о наборах требований API диалоговых окон
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7987a1617125f218ba883e834cb892fa9d5e2d9b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612117"
---
# <a name="dialog-api-requirement-sets"></a>Наборы обязательных элементов API диалоговых окон

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В приведенной ниже таблице перечислены наборы обязательных элементов Dialog API, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

|  Набор обязательных элементов  | Office 2013 для Windows\*<br>(единовременная покупка) | Office 2016 или более поздней версии в Windows\*<br>(единовременная покупка)   | Office для Windows<br>(версия, подключенная к подписке на Office 365) |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Сборка 15.0.4855.1000 или более поздняя | Сборка 16.0.4390.1000 или более поздняя | Версия 1602 (сборка 6741.0000) или более поздняя | 1.22 или более поздняя | 15.20 или более поздняя| Январь 2017 г. | Версия 1608 (сборка 7601.6800) или более поздняя|

>\*Пользователи одноразового приобретения Office могут не принять все исправления и обновления. Если да, то библиотека DLL, используемая Office для отправки отчетов о версии в пользовательском интерфейсе, может быть больше, чем перечисленные здесь версии, даже если обновленные библиотеки DLL, необходимые для поддержки DialogApi, не установлены на компьютере пользователя. Чтобы убедиться, что необходимое исправление установлено, пользователь должен перейти в список обновлений Office ([список office 2013](/officeupdates/msp-files-office-2013) или [список Office 2016](/officeupdates/msp-files-office-2016)), выполнить поиск **осфклиент-x-none**и установить указанное исправление.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>Dialog API 1.1

Dialog API 1.1 — это первая версия этого API. Дополнительные сведения об API можно найти в справочном разделе по [API диалоговых окон](/javascript/api/office/office.ui) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
