---
title: Наборы обязательных элементов для команд надстроек
description: Общие сведения о наборах требований для команд надстроек Office
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 5a979b5ca57cf1ddc8ebf021b72ca5fb8755a167
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612124"
---
# <a name="add-in-commands-requirement-sets"></a>Наборы обязательных элементов для команд надстроек

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Дополнительные сведения см. в статьях [Команды надстроек для Excel, Word и PowerPoint](../../design/add-in-commands.md) и [Команды надстроек Outlook](../../outlook/add-in-commands-for-outlook.md).

У первого выпуска команд надстроек нет соответствующего набора обязательных элементов (то есть набора обязательных элементов AddInCommands 1.0 не существует). В приведенной ниже таблице указаны ведущие приложения Office, которые поддерживают первый выпуск, а также их номера версии или сборки.  

| Выпуск   |  Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 для Windows<br>(единовременная покупка) | Office для Windows<br>(версия, подключенная к подписке на Office 365)   |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Команды надстроек (первый выпуск, набора обязательных элементов нет) | Н/Д | 16.0.4678.1000 *Поддерживается только в Outlook* | Версия 1809 (сборка 10827.20150) или более поздняя |Версия 1603 (сборка 6769.0000) или более поздняя | Н/Д | 15.33 или более поздняя версия| Январь 2016 г. |

В наборе обязательных элементов команд надстроек 1.1 появилась возможность [автоматического открытия области задач с документами](../../develop/automatically-open-a-task-pane-with-a-document.md).

В приведенной ниже таблице указаны наборы обязательных элементов команд надстроек 1.1, ведущие приложения Office, которые их поддерживают, и их номера версии или сборки.

|  Набор обязательных элементов  |  Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 для Windows<br>(единовременная покупка) | Office для Windows<br>(версия, подключенная к подписке на Office 365)   |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете  |  
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
