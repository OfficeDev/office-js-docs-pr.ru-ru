---
title: Наборы обязательных элементов для команд надстроек
description: ''
ms.date: 03/19/2019
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: a40107b968603311d3dea35cdd0d055adb14bf5a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451747"
---
# <a name="add-in-commands-requirement-sets"></a>Наборы обязательных элементов для команд надстроек

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Дополнительные сведения см. в статьях [Команды надстроек для Excel, Word и PowerPoint](/office/dev/add-ins/design/add-in-commands) и [Команды надстроек Outlook](/outlook/add-ins/add-in-commands-for-outlook).

У первого выпуска команд надстроек нет соответствующего набора обязательных элементов (то есть набора обязательных элементов AddInCommands 1.0 не существует). В приведенной ниже таблице указаны ведущие приложения Office, которые поддерживают первый выпуск, а также их номера версии или сборки.  

| Выпуск   |  Office 2013 для Windows | Office 2016 или более поздняя версия для Windows | Office 365 для Windows   |  Office 365 для iPad  |  Office 365 для Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Команды надстроек (первый выпуск, набора обязательных элементов нет) | Н/Д | 16.0.4678.1000 *Поддерживается только в Outlook* |Версия 1603 (сборка 6769.0000) или более поздняя | Н/Д | 15.33 или более поздняя версия| Январь 2016 г. |

В наборе обязательных элементов команд надстроек 1.1 появилась возможность [автоматического открытия области задач с документами](/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

В приведенной ниже таблице указаны наборы обязательных элементов команд надстроек 1.1, ведущие приложения Office, которые их поддерживают, и их номера версии или сборки.

|  Набор обязательных элементов  |  Office 2013 для Windows | Office 2016 или более поздняя версия для Windows | Office 365 для Windows   |  Office 365 для iPad  |  Office 365 для Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | Н/Д | 16.0.4678.1000 *Поддерживается только в Outlook*  | Версия 1705 (сборка 8121.1000) или более поздняя | Н/Д | 15.34 или более поздняя версия\*| Май 2017 г. |

>\*Метод [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) ошибочно возвращает значение `false` для версий 16.9&ndash;16.14 (включительно), но набор обязательных элементов *поддерживается* в этих версиях.

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
