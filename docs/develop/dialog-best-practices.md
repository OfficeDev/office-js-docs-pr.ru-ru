---
title: Рекомендации и правила Office dialog API
description: Содержит правила и практические методики для API диалогов Office, такие как лучшие методики для одношагового приложения (SPA)
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 4359d116e9720255278c5b3f543b135013c7e76c
ms.sourcegitcommit: 7cd501d0fdbbd4636bd08647b638dd5ca4c7c630
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50282984"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Рекомендации и правила Office dialog API

В этой статье данная статья содержит правила, приемы и практические методики для API диалогов Office, в том числе лучшие методики разработки пользовательского интерфейса диалогового окно и использования API в одношаговом приложении (SPA)

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования API диалогов Office, как описано в статье ["Использование dialog API Office](dialog-api-in-office-add-ins.md)в надстройки Office".
> 
> См. [также обработку ошибок и событий в диалоговом окне Office.](dialog-handle-errors-events.md)

## <a name="rules-and-gotchas"></a>Правила и подсказки

- Диалоговое окно может переходить только по URL-адресам HTTPS, а не по HTTP.
- URL-адрес, переданный [методу displayDialogAsync,](/javascript/api/office/office.ui) должен быть в том же домене, что и сама надстройка. Это не может быть поддомен. Однако передаемая на нее страница может перенаправляться на страницу в другом домене.
- В окне ведущего окна, которое может быть [](../reference/manifest/functionfile.md) файлом области задач или файлом функции без пользовательского интерфейса команды надстройки, может одновременно открываться только одно диалоговое окно.
- В диалоговом окне могут быть вызваны только два API Office:
  - Функция [messageParent.](/javascript/api/office/office.ui#messageparent-message-)
  - `Office.context.requirements.isSetSupported`(Дополнительные сведения см. в [подразделе "Указание приложений Office и требований к API".](specify-office-hosts-and-api-requirements.md)
- Функция [messageParent](/javascript/api/office/office.ui#messageparent-message-) может быть вызвана только со страницы в том же домене, что и сама надстройка.

## <a name="best-practices"></a>Рекомендации

### <a name="avoid-overusing-dialog-boxes"></a>Избегайте чрезмерного ского окна

Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий. При планировании контактной зоны помните, что в области задач можно использовать вкладки. Пример области задач с вкладками см. в примере Надстройка [Excel JavaScript SalesTracker.](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)

### <a name="designing-a-dialog-box-ui"></a>Проектирование пользовательского интерфейса диалоговых окно

Практические практики в диалоговом окне см. в диалоговых окнах [в надстройки Office.](../design/dialog-boxes.md)

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>Обработка блокировщиков всплывающих окон с помощью Office в Интернете

Попытка отобразить диалоговое окно при использовании Office в Интернете может привести к блокированию этого диалоговых окна блокатором всплывающих блоков браузера. В Office в Интернете есть функция, которая позволяет окнам надстройки быть исключением из блокатора всплывающих элементов браузера. Когда код вызывает метод, Office в Интернете откроет запрос, аналогичный `displayDialogAsync` следующему.

![Screenshot showing the prompt with a brief description and Allow and Ignore buttons that an add-in can generate to avoid in-browser pop-up blockers](../images/dialog-prompt-before-open.png)

Если пользователь **нажмет "Разрешить",** откроется диалоговое окно Office. Если пользователь нажмет **"Игнорировать",** запрос закроется, а диалоговое окно Office не откроется. Вместо этого `displayDialogAsync` метод возвращает ошибку 12009. Код должен перехватить эту ошибку и либо предоставить альтернативный интерфейс, не требующий диалоговое окно, либо показать пользователю сообщение о том, что надстройка требует от него разрешить диалоговое окно. (Подробнее об ошибке 12009 см. в [подзагонах displayDialogAsync.)](dialog-handle-errors-events.md#errors-from-displaydialogasync)

Если по какой-либо причине вы хотите отключить эту функцию, код должен отказаться от нее. Он делает этот запрос с [объектом DialogOptions,](/javascript/api/office/office.dialogoptions) который передается `displayDialogAsync` методу. В частности, объект должен включать `promptBeforeOpen: false` . Если для этого параметра установлено false, Office в Интернете не будет предложено пользователю разрешить надстройка открыть диалоговое окно, а диалоговое окно Office не откроется.

### <a name="do-not-use-the-_host_info-value"></a>Не используйте значение \_ информации \_ о хост-сайте

Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. Он будет примеен после параметров настраиваемого запроса, если таковые есть. Он не будет прибавлен к последующим URL-адресам, на которые будет перемещаться диалоговое окно. Корпорация Майкрософт может изменить содержимое этого значения или полностью удалить его, поэтому код не должен его читать. Это же значение добавляется в хранилище сеанса диалоговых окон (то есть свойство [Window.sessionStorage).](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) *Ваш код не должен ни считывать это значение, ни записывать в него данные*.

### <a name="opening-another-dialog-immediately-after-closing-one"></a>Открытие другого диалоговое окно сразу после его закрытия

На данной хост-странице нельзя открыть несколько диалогов, поэтому код должен вызывать [Dialog.close](/javascript/api/office/office.dialog#close__) в открытом диалоговом оке, прежде чем вызывать его, чтобы открыть `displayDialogAsync` другое диалоговое окно. Метод `close` является асинхронным. По этой причине при вызове сразу после вызова первое диалоговое окно может не закрыться полностью при попытке `displayDialogAsync` `close` Office открыть второе. В этом случае Office возвратит ошибку [12007:](dialog-handle-errors-events.md#12007) "Операция не удалась, так как у этой надстройки уже есть активное диалоговое окно".

Метод не принимает параметр обратного вызова и не возвращает объект Promise, поэтому его нельзя ожидать с помощью ключевого слова или `close` `await` `then` метода. По этой причине мы предлагаем следующую методику, когда вам нужно открыть новое диалоговое окно сразу после закрытия диалоговое окно: инкапсулировать код, чтобы открыть новое диалоговое окно в методе, и проектировать метод для рекурсивного вызова, если возвращается `displayDialogAsync` `12007` вызов . Ниже приведен пример.

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        openSecondDialog();
      }
      else {
         // Handle errors
      }
    }
  );
}
 
function openSecondDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
    (result) => {
      if(result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          openSecondDialog(); // Recursive call
        }
        else {
         // Handle other errors
        }
      }
    }
  );
}
```

Кроме того, можно принудительно приостановить код, прежде чем он попытается открыть второе диалоговое окно с помощью [метода setTimeout.](https://www.w3schools.com/jsref/met_win_settimeout.asp) Ниже приведен пример.

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        setTimeout(() => { 
          Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
             (result) => { /* callback body */ }
          );
        }, 1000);
      }
      else {
         // Handle errors
      }
    }
  );
}
```

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>Best practices for using the Office dialog API in an SPA

Если ваша надстройка использует клиентскую маршрутику, как это обычно делают однострковые приложения (SPAS), вы можете передать URL-адрес маршрута методу [displayDialogAsync](/javascript/api/office/office.ui) вместо URL-адреса отдельной HTML-страницы. *Мы не рекомендуем делать это по причинам, которые приведены ниже.*

> [!NOTE]
> Эта статья не относится к *серверной* маршрутике, например в веб-приложении express.

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Проблемы с SPAs и API диалогов Office

Диалоговое окно Office находится в новом окне с собственным экземпляром ямы JavaScript, поэтому это собственный контекст выполнения. Если вы передаете маршрут, ваша базовая страница и весь код инициализации и начальной загрузки снова запускаются в этом новом контексте, и всем переменным в диалоговом окне задаются исходные значения. Поэтому этот метод загружает и запускает второй экземпляр приложения в окне окна, что частично не позволяет использовать spa. Кроме того, код, который изменяет переменные в диалоговом окне, не изменяет версию одной и той же переменной области задач. Аналогично, диалоговое окно имеет собственное хранилище сеанса (свойство [Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) которое не доступно из кода в области задач. Диалоговое окно и хост-страница, на которой был вызван этот сервер, выглядят как два разных `displayDialogAsync` клиента на сервере. (Напоминание о том, что такое хост-страница, см. в диалоговом окне "Открытие диалоговых [окно" на хост-странице.)](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)

Таким образом, если вы передали маршрут методу, у вас не будет действительно spa; у вас будет два экземпляра одного и того `displayDialogAsync` *же SPA.* Кроме того, большая часть кода в экземпляре области задач никогда не будет использоваться в этом экземпляре, а большая часть кода в экземпляре диалоговых окнах никогда не будет использоваться в этом экземпляре. Это соответствует применению двух одностраничных приложений в одном пакете.

#### <a name="microsoft-recommendations"></a>Рекомендации Майкрософт

Вместо передачи клиентского маршрута методу рекомендуется сделать одно из следующих `displayDialogAsync` способов:

* Если код, который нужно выполнить в диалоговом окне, достаточно сложный, создайте два разных spAs явным образом; то есть иметь два spAs в разных папках одного домена. В диалоговом окне запускается один SPA-сайт, а другой на хост-странице `displayDialogAsync` диалогового окна, где был вызван. 
* В большинстве сценариев в диалоговом окне требуется только простая логика. В таких случаях проект будет значительно упрощен, разместив в домене вашего SPA одну HTML-страницу с внедренным кодом JavaScript или ссылкой на нее. Передайте URL-адрес страницы в метод `displayDialogAsync`. Это означает, что вы отклонился от литеральной идеи однобуквального приложения; У вас нет ни одного экземпляра SPA при использовании Dialog API Для Office.
