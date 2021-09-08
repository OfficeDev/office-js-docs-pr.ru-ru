---
title: Создание надстройки области задач Excel с помощью React
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и React.
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 74a28f3914ddbc54188d3b8baa33fc1faa7a30fe
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937190"
---
# <a name="use-react-to-build-an-excel-task-pane-add-in"></a>Создание надстройки области задач Excel с помощью React

В этой статье описывается процесс создания надстройки в области задач Excel с помощью React и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые условия

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project using React framework`
- **Выберите тип сценария:** `TypeScript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `Excel`

![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, где в качестве типа проекта установлена инфраструктура React.](../images/yo-office-excel-react-2.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простой надстройки области задач. Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже. Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.

- Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки. Дополнительные сведения о файле **manifest.xml** см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).
- В файле **./src/taskpane/taskpane.html** определена HTML-инфраструктура области задач, а файлы в папке **./src/taskpane/components** определяют разные части пользовательского интерфейса области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/components/App.tsx** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.

## <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

1. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

1. Выберите любой диапазон ячеек на листе.

1. Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.

    ![Снимок экрана: Excel с открытой областью задач надстройки и выделенной кнопкой "Выполнить".](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку панели задач Excel с помощью React! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>См. также

- [Руководство по надстройкам Excel](../tutorials/excel-tutorial.md)
- [Объектная модель JavaScript для Excel в надстройках Office](../excel/excel-add-ins-core-concepts.md)
- [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)