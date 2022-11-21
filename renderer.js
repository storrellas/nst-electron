document.getElementById('button_export').addEventListener('click', async () => {
  console.log("about to export")
  const full_path = document.getElementById('input_file').files[0].path
  const launch_result = await window.piccini.launch(full_path)
  document.getElementById('div_status').innerHTML = `Status: ${launch_result} rows were read`
  console.log("ret ", launch_result)
})
