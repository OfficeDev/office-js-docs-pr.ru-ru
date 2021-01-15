---
title: Наборы обязательных элементов API диалоговых окон
description: Узнайте больше о наборах требований Dialog API.
ms.date: 09/14/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 79b6960387519ac3c8b41b0b31cf6f40b5e7e067
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771363"
---
# <a name="dialog-api-requirement-sets"></a>Наборы обязательных элементов API диалоговых окон

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований Dialog API, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборки или версии приложения Office.

|  Набор обязательных элементов  | Office 2013 для Windows\*<br>(единовременная покупка) | Office 2016 или более поздней версии для Windows\*<br>(единовременная покупка)   | Office для Windows<br>(подписка) |  Office для iPad<br>(подписка)  |  Office для Mac<br>(подписка)  | Office в Интернете  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | Н/Д | Н/Д | См. службу поддержки<br>раздел ниже | 2.67 или более поздней | 16.37 или более поздней | Июнь 2020 г. | Недоступно |
| DialogApi 1.1  | Сборка 15.0.4855.1000 или более поздняя | Сборка 16.0.4390.1000 или более поздняя | Версия 1602 (сборка 6741.0000) или более поздняя | 1.22 или более поздняя | 15.20 или более поздняя | Январь 2017 г. | Версия 1608 (сборка 7601.6800) или более поздняя|

>\* Пользователи разовой покупки Office могут принять не все исправления и обновления. Если да, то DLL, которую Office использует для отчетов о своей версии в пользовательском интерфейсе, может быть больше, чем указанные здесь версии, даже если обновленные DLL, необходимые для поддержки DialogApi, не были установлены на компьютере пользователя. Чтобы убедиться, что установлено необходимое исправление, пользователь должен перейти в список обновлений Office (список [Office 2013](/officeupdates/msp-files-office-2013) или [Список Office 2016),](/officeupdates/msp-files-office-2016)найти **osfclient-x-none** и установить указанный исправление.

## <a name="office-on-windows-subscription-support"></a>Поддержка Office для Windows (подписка)

Набор требований DialogApi 1.2 поддерживается в канале consumer Channel версии 2005 (сборка 12827.20268 или более новой). Для Office для Windows эта функция также поддерживается в сборках канала Semi-Annual Channel и Monthly Enterprise Channel, доступных 9 июня 2020 г. или более поздней версии. Минимальные поддерживаемые сборки для каждого канала:  

|Канал | Версия | Сборка|
|:-----|:-----|:-----|
|Актуальный канал | 2005 или более | 12827.20160 или более|
|Ежемесячный канал (корпоративный) | 2004 или более | 12730.20430 или более|
|Полугодовой канал (корпоративный) | 2002 или более | 12527.20720 или более|

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11-and-12"></a>Dialog API 1.1 и 1.2

Dialog API 1.1 — это первая версия этого API. Набор требований 1.2 добавляет поддержку отправки данных с родительской страницы в диалоговое окно с помощью метода [Office.dialog.messageChild.](/javascript/api/office/office.dialog#messageChild_message_) Подробные сведения об этих API см. в справочном разделе [по Dialog API.](/javascript/api/office/office.ui)

## <a name="see-also"></a>См. также

- [Использование Office Dialog API в надстройках Office](../../develop/dialog-api-in-office-add-ins.md)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
