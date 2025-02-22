// --- Funções do Chat com persistência de histórico ---
function loadChatHistory() {
    var storedHistory = localStorage.getItem('chatHistory');
    if (storedHistory) {
      document.getElementById('chat-messages').innerHTML = storedHistory;
    }
  }

function saveChatHistory() {
    var chatContent = document.getElementById('chat-messages').innerHTML;
    localStorage.setItem('chatHistory', chatContent);
}

function clearChat() {
  document.getElementById('chat-messages').innerHTML = "";
  localStorage.removeItem('chatHistory');
  sessionStorage.removeItem("chatOpened"); // Remove a flag de chat aberto
}

window.clearChat = clearChat;

function toggleChat() {
  var chatOverlay = document.getElementById("chat-overlay");
  var chatMessages = document.getElementById("chat-messages");

  if (chatOverlay.style.display === "none" || chatOverlay.style.display === "") {
      chatOverlay.style.display = "block";

      var storedHistory = localStorage.getItem('chatHistory');

      // Se não há histórico salvo, ou o chat foi limpo, exibe mensagem de boas-vindas
      if (!storedHistory || storedHistory.trim() === "") {
          chatMessages.innerHTML = ""; // Limpa mensagens antigas

          var botDiv1 = document.createElement("div");
          botDiv1.classList.add("chat-message", "bot");
          botDiv1.textContent = "Bem-vindo ao chatbot! Envie uma mensagem para começar.";

          var botDiv2 = document.createElement("div");
          botDiv2.classList.add("chat-message", "bot");
          botDiv2.textContent = "Se precisar de ajuda, basta perguntar!";

          var botDiv3 = document.createElement("div");
          botDiv3.classList.add("chat-message", "bot");
          botDiv3.innerHTML = `
            <strong>Sugestões de perguntas:</strong><br>
            <span class="highlight-grey">! Apenas trabalho com períodos de MM/AAAA, ou só o mês (seja o mês por nome ou numeral) e/ou ano completo.</span><br><br>
            • Listagem de nomes<br>
            • Nome com mais <span class="highlight-blue">[tipo de presença]</span> em <span class="highlight-blue">[período]</span><br>
            • Presença geral em <span class="highlight-blue">[período]</span><br>
            • <span class="highlight-blue">[Nome]</span> teve presença em <span class="highlight-blue">[período]</span><br>
            • Quantas <span class="highlight-blue">[tipo de presença]</span> teve <span class="highlight-blue">[Nome]</span> em <span class="highlight-blue">[período]</span><br>
        `;
      


          chatMessages.appendChild(botDiv1);
          chatMessages.appendChild(botDiv2);
          chatMessages.appendChild(botDiv3);

          saveChatHistory();
      } else {
          chatMessages.innerHTML = storedHistory; // Carrega o histórico se existir
      }

      sessionStorage.setItem("chatOpened", "true"); // Marca que o chat foi aberto
  } else {
      chatOverlay.style.display = "none";
  }
}

window.toggleChat = toggleChat;

function handleKeyPress(event) {
    if (event.key === "Enter") {
      sendMessage();
    }
}

window.handleKeyPress = handleKeyPress;

function sendMessage() {
    var inputField = document.getElementById("chat-input");
    var userMessage = inputField.value.trim();
    if (userMessage === "") return;

    var chatMessages = document.getElementById("chat-messages");

    // Cria e adiciona o elemento para a mensagem do usuário
    var userDiv = document.createElement("div");
    userDiv.classList.add("chat-message", "user");
    userDiv.textContent = userMessage;
    chatMessages.appendChild(userDiv);
    saveChatHistory();

    // Envia a mensagem para o backend (rota /chatbot)
    fetch("/chatbot", {
      method: "POST",
      body: JSON.stringify({ mensagem: userMessage }),
      headers: { "Content-Type": "application/json" }
    })
      .then(response => response.json())
      .then(data => {
        if (data.respostas) {
          data.respostas.forEach(function (resp) {
            var botDiv = document.createElement("div");
            botDiv.classList.add("chat-message", "bot");
            // Se o tipo for "table", insere como HTML para renderizar a tabela;
            // caso contrário, insere como texto.
            if (resp.tipo === "table") {
              botDiv.innerHTML = resp.mensagem;
            } else {
              botDiv.textContent = resp.mensagem;
            }
            chatMessages.appendChild(botDiv);
          });
        }
        chatMessages.scrollTop = chatMessages.scrollHeight;
        saveChatHistory();
      })
      .catch(error => console.error("Erro ao enviar mensagem:", error));

    inputField.value = "";
}
window.sendMessage = sendMessage;

loadChatHistory();