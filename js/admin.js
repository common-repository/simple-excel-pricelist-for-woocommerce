document.addEventListener("DOMContentLoaded", function(){

  var gen_btn = document.getElementById('sepw-generate');
  var ok_message = document.getElementById('sepw-generate-ok');
  var error_message = document.getElementById('sepw-generate-error');
  var generated_time = document.getElementById('sepw-generated-time');

  if (gen_btn) {
    gen_btn.addEventListener('click', generate_click);
  }

  function generate_click(event) {
    event.preventDefault();

    gen_btn.setAttribute('disabled', true);
    ok_message.style.display = 'none';
    error_message.style.display = 'none';

    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = ready_callback;
    xhttp.open("GET", "/wp-json/sepw/v1/generate", true);
    xhttp.send();
  }

  function ready_callback() {
    if (this.readyState != 4) return;
    if (this.status == 200) {
      var json = this.responseText;
      var response = {};
      try {
        response = JSON.parse(json);
      } catch(e) {};
      ok_message.style.display = 'block';
      generated_time.innerHTML = response.time;
    } else {
      error_message.style.display = 'block';
    }
    gen_btn.removeAttribute('disabled');
  }

});
