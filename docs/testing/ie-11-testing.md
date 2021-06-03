---
title: Тестирование Internet Explorer 11
description: Проверьте Office надстройки в Internet Explorer 11.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: de256ee8b0633f18d3188c5bbfae52cb24ff2c35
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727936"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Проверьте Office надстройки в Internet Explorer 11

Если вы планируете выставлять надстройку на рынок через AppSource или планируете поддерживать более старые версии Windows и Office, надстройка должна работать в встраиваемом контроле браузера, основанном на Internet Explorer 11 (IE11). Вы можете использовать командную строку для перехода от более современных времен работы, используемых надстройки, к времени запуска Internet Explorer 11 для этого тестирования. Сведения о том, какие версии Windows и Office используют управление веб-представлением Internet Explorer 11, см. в браузерах, используемых Office [надстройки.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней части, у вас есть два варианта:
>
> - Напишите код в ECMAScript 2015 (также называемый ES6) или позже JavaScript, или в TypeScript, а затем скомпилировать код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
> - Напишите в ECMAScript 2015 или более [](https://en.wikipedia.org/wiki/Polyfill_(programming)) поздний JavaScript, а также загрузите библиотеку полифильмов, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать код.
>
> Дополнительные сведения об этих параметрах см. в [меню Support Internet Explorer 11.](../develop/support-ie-11.md)
>
> Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.

> [!NOTE]
> Чтобы протестировать надстройку в браузере Internet Explorer 11, откройте Office в Интернете в Internet Explorer и разгрузите [надстройку.](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))

Эти инструкции предполагают, что вы создали проект генератора Yo Office ранее. Если вы еще не сделали этого раньше, рассмотрите возможность быстрого начала чтения, например для Excel [надстройки.](../quickstarts/excel-quickstart-jquery.md)

## <a name="switching-to-the-internet-explorer-11-webview"></a>Переход на веб-просмотр Internet Explorer 11

1. Создайте проект Office Yo. Неважно, какой проект вы выберете, этот инструментарий будет работать со всеми типами проектов.

    > [!NOTE]
    > Если у вас есть существующий проект и вы хотите добавить этот инструмент без создания нового проекта, пропустите этот шаг и перейдйте к следующему шагу. 

1. В корневой папке проекта запустите следующую строку в командной строке. В этом примере предполагается, что файл манифеста проекта находится в корне. Если это не так, укажите относительный путь к файлу манифеста. В командной строке должно быть видно сообщение о том, что тип веб-представления теперь настроен на IE.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> Эта команда не требуется, но она должна помочь отламеть большинство проблем, связанных с запуском Internet Explorer 11. Для полной надежности необходимо проверить использование компьютеров с различными комбинациями Windows 7, 8.1 и 10 и различных Office. Дополнительные сведения [](../concepts/browsers-used-by-office-web-add-ins.md) см. в Office надстройки и сведения о том, как вернуться к более ранней версии [Office.](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)

### <a name="command-options"></a>Параметры команды

В качестве аргументов команда может также использовать несколько времен `office-addin-dev-settings webview` работы:

- ie
- edge
- default

## <a name="see-also"></a>См. также

* [Тестирование и отладка надстроек Office](test-debug-office-add-ins.md)
* [Загрузка неопубликованных надстроек Office для тестирования](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Отладка надстроек с помощью средств разработчика в Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Подключение отладчика из области задач](attach-debugger-from-task-pane.md)
