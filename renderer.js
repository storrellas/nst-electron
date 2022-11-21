document.getElementById('export').addEventListener('click', async () => {
  console.log("about to export")
  const launch_result = await window.piccini.launch()
  console.log("ret ", launch_result)
  // const isDarkMode = await window.darkMode.toggle()
  // document.getElementById('theme-source').innerHTML = isDarkMode ? 'Dark' : 'Light'
})
