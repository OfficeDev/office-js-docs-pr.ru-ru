
# <a name="walkthrough-creating-your-first-lab-for-office-mix"></a>Создание первой лаборатории для Office Mix — пошаговое руководство
Создание первой лаборатории LabsJS с помощью с пошагового руководства.



Это пошаговое руководство поможет вам создать с нуля простую лабораторию LabsJS. Это будет простой тест "правда/неправда" с одним вопросом. 

Вы начнете не с шаблона проекта Visual Studio, а с трех пустых файлов, что свидетельствует о простоте лаборатории: 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
Для редактирования этих файлов подходит любой редактор кода, потому что мы не используем вначале шаблон Visual Studio. На самом деле HTML-файл очень прост, и при желании вы можете просто скопировать и вставить разметку HTML из учебных файлов. Но обратите внимание, что требуется формат HTML5, поэтому объявлять тип документа нужно так: `<!DOCTYPE html>`. CSS-файл не является обязательным. Основная часть работы выполняется в файле JavaScript (JS), TrueFalse.js. В этом руководстве будут рассмотрены четыре основные функции лаборатории:

- установка (подключение к узлу);
    
- смена режимов (редактирования и просмотра);
    
- изменение лаборатории;
    
- выполнение (или запуск) лаборатории.
    

 **Примечание**  
 ---
 Файл labhost.html запускается на веб-сервере и обеспечивает среду внешнего размещения для разработки и тестирования лаборатории. Это значительно упрощает разработку лаборатории. Сведения о настройке среды разработки см. в статье [Начало работы с LabsJS для Office Mix](get-started-with-labsjs-for-office-mix.md).<br/><br/>

И, наконец, можно увидеть выполненный файл JavaScript (TrueFalse.js) среди файлов, распространенных с помощью данного пакета SDK. Ниже приведено пошаговое руководство по процессу кодирования.

## <a name="connecting-to-the-lab-host"></a>Подключение лаборатории к узлу

Лаборатории в этой среде могут работать как с нашим узлом (для разработки и тестирования), так и с узлом среды выполнения по умолчанию — узлом Office.js. Затем функция открытия использует простое выражение if/else, чтобы определить, какой из этих контекстов размещения подходит.


```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

Объект **PostMessageLabHost** запускается в среде разработки labhost.html, тогда как в рабочей среде лаборатория запускается в PowerPoint/Office Mix с помощью **OfficeJSLabHost**.

Затем нужно создать вспомогательный метод для выполнения обратного вызова, который должен разрешить или отклонить передаваемый отложенный объект jQuery. Используйте метод **createCallback**, чтобы перейти от обещаний jQuery к обратным вызовам, которые определяются файлом labs.js.




```js
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

Кроме того, мы создаем вспомогательный метод получения конфигурации лаборатории для конкретного вопроса и ответа.




```js
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## <a name="mode-changes"></a>Смена режимов

Лаборатория всегда находится в одном из двух состояний, или режимов: **view** (просмотр) и **edit** (редактирование). Поэтому для теста нам нужен способ захвата и удержания состояния (режима). С этой целью мы создаем класс.


```js
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

Кроме того, нам нужно создать метод обновления пользовательского интерфейса теста, где ответ на вопрос (т. е. его отправка) является правильным или неправильным.




```js
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

Нам также нужна функция, позволяющая переключаться между режимами редактирования и просмотра.




```js
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

Наша следующая функция обновляет конфигурацию теста, основанного на событиях изменения, полученных из пользовательского интерфейса.




```js
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

Затем мы обновляем данные конфигурации лаборатории, которые хранятся на сервере и зависят от конкретного вопроса и ответа.




```js
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

Кроме того, мы используем функцию, которая связывает обновления конфигурации, выполненные в лаборатории в режиме редактирования. За ней следует код для отмены имеющейся привязки к обработчикам изменений.




```js
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```js
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

Теперь приступим к ключевой части раздела, а именно к методам переключения между режимами просмотра и редактирования. Начнем с перехода из режима просмотра в режим редактирования.




```js
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

А теперь посмотрим, как перейти из режима редактирования в режим просмотра.




```js
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

И, наконец, после подключения к узлу и подготовки документа можно запускать тест.




```js
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## <a name="additional-resources"></a>Дополнительные ресурсы
<a name="bk_addresources"> </a>


- [Надстройки Office Mix](office-mix-add-ins.md)
    
