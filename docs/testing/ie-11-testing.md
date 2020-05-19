---
ms.date: 05/16/2020
description: Протестируйте надстройку Office с помощью Internet Explorer 11.
title: Тестирование Internet Explorer 11
localization_priority: Normal
ms.openlocfilehash: 697c87d90df9aa70a7b20da5cd4c91d4445fb850
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275950"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a>Тестирование надстройки Office с помощью Internet Explorer 11

В зависимости от спецификаций надстройки вы можете запланировать поддержку более ранних версий Windows и Office, которые требуют тестирования в Internet Explorer 11. Это часто требуется при отправке надстройки в AppSource. С помощью средства командной строки можно переключиться с более современных сред выполнения, используемых надстройками, в среду выполнения Internet Explorer 11 для этого тестирования.

## <a name="pre-requisites"></a>Необходимые условия

- [Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))
- Редактор кода. Мы рекомендуем [Visual Studio Code](https://code.visualstudio.com/)
- [Участие в программе предварительной оценки Office](https://insider.office.com)

В этих инструкциях предполагается, что ранее был настроен проект генератора Yo Office. Если вы еще этого не сделали, рекомендуем ознакомиться со кратким руководством, например: [для надстроек Excel](../quickstarts/excel-quickstart-jquery.md).

## <a name="using-ie11-tooling"></a>Использование средства IE11

1. Создайте проект генератора Yo Office. В этом случае не имеет значения, какой тип проекта будет выбран, это средство будет работать со всеми типами проектов.

> ! НОТЕ Если у вас есть проект и вы хотите добавить этот инструмент без создания нового проекта, пропустите этот шаг и перейдите к следующему шагу. 

2. В корневой папке нового проекта выполните в командной строке следующую команду:

```command&nbsp;line
office-add-dev-settings webview manifest.xml ie
```
В командной строке должно появиться примечание о том, что в качестве типа представления веб-сайта теперь задано значение IE.

> ! Последняя Это средство не обязательно использовать, но оно должно помочь отладить большинство проблем, связанных со средой выполнения Internet Explorer 11. Для полной надежности необходимо протестировать использование компьютера с установленной копией Windows 7 и Office 2013.

## <a name="command-settings"></a>Параметры команды

Если у вас есть другой путь манифеста, укажите его в команде, как показано в следующем примере:

`office-add-dev-settings webview [path to your manifest] ie`

`office-addin-dev-settings webview`Кроме того, в качестве аргументов команды можно использовать ряд сред выполнения:

- Explorer
- кромки
- умолчани

## <a name="see-also"></a>См. также
* [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)
* [Загрузка неопубликованных надстроек Office для тестирования](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Отладка надстроек с помощью средств разработчика в Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Подключение отладчика из области задач](attach-debugger-from-task-pane.md)