---
title: Наборы обязательных элементов для команд надстроек
description: Общие сведения о наборах требований для команд надстроек Office.
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 25c8a08983d617e4592dd5602d06eb1d780165d0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293592"
---
# <a name="add-in-commands-requirement-sets"></a>Наборы обязательных элементов для команд надстроек

Наборы требований — это именованные группы элементов API. Надстройки Office используют наборы требований, указанные в манифесте, или используют проверку среды выполнения, чтобы определить, поддерживает ли приложение Office API, необходимые надстройке. Более подробную информацию можно узнать в статье [версии Office и наборах требований](../../develop/office-versions-and-requirement-sets.md).

Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Дополнительные сведения см. в статьях [Команды надстроек для Excel, Word и PowerPoint](../../design/add-in-commands.md) и [Команды надстроек Outlook](../../outlook/add-in-commands-for-outlook.md).

В первом выпуске команд надстроек отсутствует соответствующий набор обязательных элементов (то есть набор требований Аддинкоммандс 1,0 не существует). В следующей таблице перечислены клиентские приложения Office, поддерживающие первоначальную версию выпуска, а также номера версий или номера сборок для этих приложений.  

| Выпуск   |  Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365)   |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Команды надстроек (первый выпуск, набора обязательных элементов нет) | Н/Д | 16.0.4678.1000 *Поддерживается только в Outlook* | Версия 1809 (сборка 10827.20150) или более поздняя |Версия 1603 (сборка 6769.0000) или более поздняя | Н/Д | 15.33 или более поздняя версия| Январь 2016 г. |

В наборе обязательных элементов команд надстроек 1.1 появилась возможность [автоматического открытия области задач с документами](../../develop/automatically-open-a-task-pane-with-a-document.md).

В следующей таблице перечислены наборы обязательных элементов для команд надстроек 1,1, клиентские приложения Office, поддерживающие этот набор требований, а также номера сборок или версий приложений Office.

|  Набор обязательных элементов  |  Office 2013 для Windows<br>(единовременная покупка) | Office 2016 для Windows<br>(единовременная покупка) | Office 2019 для Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365)   |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |  
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
- [Указание приложений Office и требований к API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
