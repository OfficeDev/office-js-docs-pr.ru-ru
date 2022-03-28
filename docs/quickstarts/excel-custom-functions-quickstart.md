---
ms.date: 03/23/2022
description: Краткое руководство по разработке пользовательских функций в Excel.
title: Краткое руководство по пользовательским функциям
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: cac81cb25b9880a3057e2246d39ac226666a4cb4
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404710"
---
# <a name="get-started-developing-excel-custom-functions"></a>Начало разработки пользовательских функций Excel

С помощью пользовательских функций разработчики могут добавлять новые функции в Excel, определяя их в JavaScript или Typescript как часть надстройки. Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.

## <a name="prerequisites"></a>Предварительные требования

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Пакет Office, подключенный к подписке Microsoft 365 (включая Office в Интернете).

  > [!NOTE]
  > Если у вас еще нет Office, вы можете [присоединиться к программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program), чтобы получить бесплатную 90-дневную возобновляемую подписку на Microsoft 365 для использования в процессе разработки.

## <a name="build-your-first-custom-functions-project"></a>Создание первого проекта пользовательских функций

Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman. Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Выберите тип проекта:** `Excel Custom Functions Add-in project`
    - **Выберите тип сценария:** `JavaScript`
    - **Как вы хотите назвать надстройку?** `starcount`

    :::image type="content" source="../images/starcountPrompt.png" alt-text="Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, запрашивающий проекты пользовательских функций.":::

    Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.

1. Генератор Yeoman предоставит вам инструкции в командной строке по действиям с проектом, но вам нужно их проигнорировать и продолжить выполнять наши инструкции. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd starcount
    ```

1. Выполните построение проекта.

    ```command&nbsp;line
    npm run build
    ```

1. Запустите локальный веб-сервер, работающий на Node.js. Вы можете попробовать использовать надстройку пользовательской функции в Excel. Вам может быть предложено открыть область задач надстройки, но это необязательно. Вы по-прежнему можете запускать свои пользовательские функции, не открывая область задач надстройки.

# <a name="excel-on-windows-or-mac"></a>[Excel для Windows или Mac](#tab/excel-windows)

Чтобы проверить надстройку в Excel для Windows или Mac, выполните следующую команду. Когда вы выполните эту команду, запустится локальный веб-сервер и откроется приложение Excel, в котором будет загружена ваша надстройка.

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# <a name="excel-on-the-web"></a>[Excel в Интернете](#tab/excel-online)

Чтобы проверить надстройку в Excel в Интернете, выполните следующую команду. После выполнения этой команды запустится локальный веб-сервер. Замените "{url}" на URL-адрес документа Excel в OneDrive или библиотеке SharePoint, для которой у вас есть разрешения.

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## <a name="try-out-a-prebuilt-custom-function"></a>Проверка работы готовой пользовательской функции

Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые готовые пользовательские функции, определенные в файле **./src/functions/functions.js**. Файл **./manifest.xml** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.

В книге Excel проверьте, как работает пользовательская функция `ADD`, выполнив описанные ниже шаги.

1. Выделите ячейку и введите `=CONTOSO`. Обратите внимание, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.

1. Запустите функцию `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.

Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров. При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel! Затем создайте более сложную надстройку с возможностью потоковой передачи данных. Следующая ссылка поможет вам выполнить дальнейшие действия в руководстве по надстройке Excel с пользовательскими функциями.

> [!div class="nextstepaction"]
> [Руководство по надстройке Excel для пользовательских функций](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="troubleshooting"></a>Устранение неполадок

При многократном выполнении быстрого запуска могут возникнуть проблемы. Если в кэше Office уже есть экземпляр функции с таким же именем, в вашей надстройке возникнет ошибка при ее загрузке без публикации. Это можно предотвратить путем [очистки кэша Office](../testing/clear-cache.md) перед запуском `npm run start`.

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="Сообщение об ошибке в Excel под названием &quot;Ошибка при установке функций&quot;. Она содержит текст &quot;Эта надстройка не была установлена, так как пользовательская функция с тем же именем уже существует&quot;.":::

## <a name="see-also"></a>См. также

- [Обзор пользовательских функций](../excel/custom-functions-overview.md)
- [Метаданные пользовательских функций](../excel/custom-functions-json.md)
- [Среда выполнения для пользовательских функций Excel](../excel/custom-functions-runtime.md)
