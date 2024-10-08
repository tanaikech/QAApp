function javascript_() {

  function manageChunks() {
    run_("manageChunks");
  }

  function clearCorpus() {
    run_("clearCorpus");
  }

  function getChunksAndUpdateDataSheet() {
    run_("getChunksAndUpdateDataSheet");
  }

  function generateAnswer() {
    const question = document.getElementById("question");
    const answer = document.getElementById("answer");
    const button = document.getElementById("generateAnswer");
    const progress = button.parentNode.querySelector("#progress");

    progress.style.display = "block";
    answer.value = "";
    M.textareaAutoResize(answer);
    const arg = question.value;
    button.classList.add("disabled");
    google.script.run.withSuccessHandler(value => {
      console.log(value);
      answer.value = value;
      M.textareaAutoResize(answer);
      button.classList.remove("disabled");
      progress.style.display = "none";
    }).generateAnswer(arg);
  }

  function run_(e) {
    const button = document.getElementById(e);
    const progress = button.parentNode.querySelector("#progress");
    progress.style.display = "block";
    button.classList.add("disabled");
    google.script.run.withSuccessHandler(_ => {
      button.classList.remove("disabled");
      progress.style.display = "none";
    })[e]();
  }

}
