---
title: Наборы обязательных элементов API диалоговых окон
description: Дополнительные информацию о наборах требований к API диалогов.
ms.date: 10/05/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 4802189b0dbde30d0d9058b542c35cac47074998
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138557"
---
# <a name="dialog-api-requirement-sets"></a>Наборы обязательных элементов API диалоговых окон

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API диалогов, Office клиентские приложения, поддерживают этот набор требований, а также номера сборки или версии для Office приложения.

| Набор обязательных элементов | Office 2013 для Windows\*<br>(единовременная покупка) | Office 2016 для Windows\*<br>(единовременная покупка) | Office 2019 для Windows\*<br>(единовременная покупка) | Office 2021 или более поздней Windows\*<br>(единовременная покупка) | Office для Windows<br>(подписка) | Office для iPad<br>(подписка) |  Office для Mac<br>(подписка) | Office в Интернете | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | Н/Д | Н/Д | Н/Д | Сборка 16.0.14326.20454 или более поздней | См. поддержку<br>раздел ниже | 2.37 или более поздней | 16.37 или более поздней | Июнь 2020 г. | Недоступно |
| DialogApi 1.1  | Сборка 15.0.4855.1000 или более поздняя | Сборка 16.0.4390.1000 или более поздняя | Сборка 16.0.12527.20720 или более поздней | Сборка 16.0.14326.20454 или более поздней | Версия 1602 (сборка 6741.0000) или более поздняя | 1.22 или более поздняя | 15.20 или более поздняя | Январь 2017 г. | Версия 1608 (сборка 7601.6800) или более поздняя|

>\*Пользователи разовой системы Office, возможно, не приняли все исправления и обновления. Если это так, то DLL, Office для отчета о своей версии в пользовательском интерфейсе, может быть больше, чем указанные здесь версии, даже если обновленные DLLs, необходимые для поддержки DialogApi, не установлены на компьютере пользователя. Чтобы обеспечить установку необходимого исправления, пользователю необходимо перейти в список обновлений Office [(Office 2013](/officeupdates/msp-files-office-2013) или [Office 2016](/officeupdates/msp-files-office-2016)г.), **поискать osfclient-x-none** и установить указанный патч.

## <a name="office-on-windows-subscription-support"></a>Office поддержки Windows (подписка)

Набор требований DialogApi 1.2 поддерживается в версии consumer Channel 2005 (сборка 12827.20268 или более). Для Office на Windows, функция также поддерживается в сборках Semi-Annual channel и Monthly Enterprise Channel, доступных 9 июня 2020 г. или более поздней периодии. Минимальные поддерживаемые сборки для каждого канала:  

|Канал | Версия | Сборка|
|:-----|:-----|:-----|
|Актуальный канал | 2005 или больше | 12827.20160 или больше|
|Ежемесячный канал (корпоративный) | 2004 или более | 12730.20430 или более|
|Полугодовой канал (корпоративный) | 2002 или больше | 12527.20720 или более|

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11-and-12"></a>API диалогов 1.1 и 1.2

Dialog API 1.1 — это первая версия этого API. Набор требований 1.2 добавляет поддержку для отправки данных с родительской страницы в диалоговое окно [методом Office.dialog.messageChild.](/javascript/api/office/office.dialog#messageChild_message_) Подробные сведения об этих API см. в справочной теме [API](/javascript/api/office/office.ui) диалогов.

## <a name="see-also"></a>См. также

- [Использование Office Dialog API в надстройках Office](../../develop/dialog-api-in-office-add-ins.md)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
