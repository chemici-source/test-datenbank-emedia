(function () {
  function init() {
    var table = document.querySelector("table.data");
    if (!table) return;

    function markTruncatable() {
      table.querySelectorAll("td:not(.truncatable):not(.expanded)").forEach(function (td) {
        if (td.scrollWidth > td.clientWidth + 1) {
          td.classList.add("truncatable");
        }
      });
    }

    markTruncatable();
    window.addEventListener("resize", markTruncatable);

    table.addEventListener("click", function (e) {
      var td = e.target.closest("td");
      if (!td || !table.contains(td)) return;
      if (!td.classList.contains("truncatable") && !td.classList.contains("expanded")) return;

      // don't collapse the cell out from under a text selection (e.g. copying)
      var sel = window.getSelection();
      if (sel && sel.toString().length > 0) return;

      td.classList.toggle("expanded");
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
