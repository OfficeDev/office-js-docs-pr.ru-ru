
# <a name="configuring-and-editing-labsjs-labs-for-office-mix"></a>Настройка и редактирование лабораторий LabsJS для Office Mix



Office Mix предоставляет методы office.js для получения и задания конфигураций лаборатории. Конфигурация указывает Office Mix тип создаваемой лаборатории, а также тип данных, которые лаборатория будет возвращать. Эти сведения используются для сбора и визуализации аналитики.

## <a name="getting-the-lab-editor"></a>Получение редактора лаборатории

Редактор лаборатории, объект [Labs.LabEditor](http://dev.office.com/reference/add-ins/office-mix/labs.labeditor), позволяет изменять лабораторию, а также получать и задавать ее конфигурацию. Завершив редактирование лаборатории, необходимо вызвать метод **Done**. Однако вызов метода **Done** требуется, только если вы пытаетесь выполнить или запустить редактируемую лабораторию. Обратите внимание, что одновременно можно открыть только один экземпляр лаборатории.

В следующем коде показано, как получить редактор лаборатории.




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

Для сохранения конфигурации конкретной лаборатории используйте методы **getConfiguration** и **setConfiguration** в [Labs.LabEditor](http://dev.office.com/reference/add-ins/office-mix/labs.labeditor). Конфигурация ([Labs.Core.IConfiguration](http://dev.office.com/reference/add-ins/office-mix/labs.core.iconfiguration)) указывает Office Mix тип данных, которые будут собираться и обрабатываться лабораторией. Конфигурация содержит общие сведения о лаборатории, включая имя, версию и другие параметры. Самая важная часть конфигурации — это определение компонентов лаборатории.

В следующем коде показано, как задать и получить конфигурацию. Чтобы задать конфигурацию, просто создайте ее объект, а затем вызовите метод **setConfiguration**. Чтобы получить конфигурацию, вызовите метод **getConfiguration** в объекте редактора лаборатории.




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## <a name="closing-the-editor"></a>Закрытие редактора

Чтобы закрыть редактор, вызовите в нем метод **Done** после того, как завершите редактирование лаборатории. Обратите внимание, что невозможно одновременно выполнять и редактировать лабораторию. Но метод **Done** позволяет либо редактировать, либо запустить лабораторию.


## <a name="interacting-with-a-lab"></a>Взаимодействие с лабораторией

Задав конфигурацию лаборатории, можно начинать взаимодействовать с ней. При запуске лаборатории в PowerPoint происходит имитация взаимодействия. Но если лаборатория запускается в проигрывателе уроков Office Mix, данные сохраняются в базе данных Office Mix и используются для аналитики.


### <a name="getting-the-lab-instance"></a>Получение экземпляра лаборатории

Взаимодействие с лабораторией происходит при помощи объекта [Labs.LabInstance](http://dev.office.com/reference/add-ins/office-mix/labs.labinstance), который является экземпляром лаборатории, настроенной для текущего пользователя. Для запуска (или выполнения) лаборатории необходимо вызвать функцию [Labs.takeLab](http://dev.office.com/reference/add-ins/office-mix/labs.takelab).


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

Экземпляр объекта содержит массив экземпляров компонентов ([Labs.ComponentInstanceBase](http://dev.office.com/reference/add-ins/office-mix/labs.componentinstancebase), [Labs.ComponentInstance](http://dev.office.com/reference/add-ins/office-mix/labs.componentinstance)), которые приводятся в соответствие с компонентами, заданными в конфигурации. По сути, экземпляр — это преобразованная версия конфигурации, которая используется для присвоения идентификаторов серверной части экземплярам объектов, а также для скрытия от пользователя определенных полей, когда это необходимо (например, подсказки, ответы и т. д.).


### <a name="managing-state"></a>Управление состоянием

Состояние — это временное хранилище, связанное с пользователем, запускающим лабораторию. Это хранилище можно использовать для сохранения данных между последовательными вызовами лаборатории. Например, лаборатория для программирования может сохранять работу пользователя в текущем состоянии.

Чтобы задать (**set**) состояние, используйте следующий код:




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

Чтобы получить (**get**) состояние, используйте такой код:




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## <a name="component-instances-and-results"></a>Экземпляры компонентов и результаты

Ниже описывается применение экземпляров четырех типов компонентов, а также приводятся краткие примеры методов компонентов. 

Но сначала следует ознакомиться с двумя основными понятиями, необходимыми при работе с экземплярами компонентов. Эти понятия — **attempts** (попытки) и **values** (значения).

 **Attempts**

Attempt — это попытка пользователя выполнить экземпляр компонента. Например, в случае вопроса с несколькими вариантами ответа попытка начинается в тот момент, когда пользователь начинает работать над проблемой, и заканчивается, когда назначается окончательная оценка. Аналитика Office Mix затем суммирует результаты пользователя, полученные при работе над проблемой.


 >**Примечание.** Попытки можно использовать для всех типов компонентов, кроме **DynamicComponent**.

Получить результаты всех попыток, связанных с конкретным экземпляром компонента, можно с помощью метода **getAttempts**. Получив результаты, пользователь может либо повторно предпринять одну из имеющихся попыток при помощи метода **resume**, либо создать другую попытку, используя метод **createAttempt**. Этот процесс показан в следующем примере:




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **Values**

Экземпляры компонентов содержат словарь ключей, которые приводятся в соответствие с массивом значений. Этот массив можно использовать для хранения подсказок, комментариев или любого другого набора значений, которые нужно связать с компонентом. Экземпляр компонента обеспечивает доступ к этим значениям при помощи метода **getValues**.

Например, в результате запроса значения подсказки аналитика отмечает, что пользователь использовал подсказку. Значения отслеживаются для каждой отдельной попытки.

В следующем примере кода показано, как запросить подсказку:




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### <a name="activitycomponentinstance"></a>ActivityComponentInstance


Объект **ActivityComponentInstace** используется для отслеживания действий пользователя с компонентом действия. В этом классе используется метод **complete**, означающий, что взаимодействие с действием завершено. Этот метод указывает, что пользователь завершил заданную задачу, закончил читать или достиг любой другой конечной точки, связанной с действием. В следующем коде показано, как использовать метод **complete**:


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### <a name="choicecomponentinstance"></a>ChoiceComponentInstance


Объект **ChoiceComponentInstance** используется для отслеживания действий пользователя с компонентом выбора. Компоненты выбора — это проблемы, которые предлагают пользователю список вариантов ответа для выбора. Правильный вариант может отсутствовать. В этом классе используются два основных метода: **getSubmissions** и **submit**. Метод **getSubmissions** позволяет получать ранее сохраненные ответы, а метод **submit** — сохранять новый ответ. Использование этих методов проиллюстрировано в следующем примере кода:


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="inputcomponentinstance"></a>InputComponentInstance


Объект **InputComponentInstance** используется для отслеживания действий пользователя с компонентом ввода. В этом классе используются два основных метода: **getSubmission** и **submit**. Метод **getSubmissions** позволяет получать ранее сохраненные ответы, а метод **submit** — сохранять новый ответ. В следующем фрагменте кода проиллюстрировано использование метода **getSubmissions**:


```js
var submissions = this._attempt.getSubmissions();
```

При использовании метода **submit** следует обратить внимание, что объект **InputComponentAnswer** представляет отправленный ответ, а объект **InputComponentResult** содержит результат. Возвращаемое значение — объект **InputComponentSubmission**, который содержит ответ, результат и метку времени отправки результата.




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="dynamiccomponentinstance"></a>DynamicComponentInstance


Объект **DynamicComponentInstance** используется для отслеживания действий пользователя с динамическим компонентом. Основные методы в этом классе: **getComponents**, **createComponent** и **close**.

Метод **getComponents** позволяет получать список ранее созданных экземпляров компонентов (см. пример ниже).




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

Метод **createComponent** создает компонент и возвращает его экземпляр, как показано в примере ниже.




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

Метод **close** указывает, что использование динамического компонента для создания компонентов завершено. Обратите внимание, что можно также использовать логический метод **isClosed**, позволяющий определить, закрыт ли экземпляр динамического компонента. Использование метода **close** показано в следующем примере:




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## <a name="additional-resources"></a>Дополнительные ресурсы



- [Надстройки Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Пошаговое руководство. Создание первой лаборатории для Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
