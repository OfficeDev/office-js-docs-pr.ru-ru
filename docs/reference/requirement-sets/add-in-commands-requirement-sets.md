---
title: Наборы обязательных элементов для команд надстроек
description: Общие сведения о наборах требований для команд надстроек Office
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d1904d092988a445be3e481123eecbad39097764
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094416"
---
# <a name="add-in-commands-requirement-sets"></a>Наборы обязательных элементов для команд надстроек

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../design/add-in-commands.md) and [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| Выпуск   |  Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке Microsoft 365)   |  Office для iPad<br>(подключено к подписке Microsoft 365)  |  Office для Mac<br>(подключено к подписке Microsoft 365)  | Office в Интернете  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Команды надстроек (первый выпуск, набора обязательных элементов нет) | Недоступно | 16.0.4678.1000 *Поддерживается только в Outlook* | Версия 1809 (сборка 10827.20150) или более поздняя |Версия 1603 (сборка 6769.0000) или более поздняя | Н/Д | 15.33 или более поздняя версия| Январь 2016 г. |

В наборе обязательных элементов команд надстроек 1.1 появилась возможность [автоматического открытия области задач с документами](../../develop/automatically-open-a-task-pane-with-a-document.md).

В приведенной ниже таблице указаны наборы обязательных элементов команд надстроек 1.1, ведущие приложения Office, которые их поддерживают, и их номера версии или сборки.

|  Набор обязательных элементов  |  Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке Microsoft 365)   |  Office для iPad<br>(подключено к подписке Microsoft 365)  |  Office для Mac<br>(подключено к подписке Microsoft 365)  | Office в Интернете  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | Н/Д | 16.0.4678.1000 *Поддерживается только в Outlook*  | Версия 1809 (сборка 10827.20150) или более поздняя | Версия 1705 (сборка 8121.1000) или более поздняя | Н/Д | 15.34 или более поздняя версия\*| Май 2017 г. |

>\*Метод [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) ошибочно возвращает значение `false` для версий 16.9&ndash;16.14 (включительно), но набор обязательных элементов *поддерживается* в этих версиях.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
