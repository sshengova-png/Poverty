const curatorNotes = [
  "Бележка от куратора: Крушата настоява да бъде наричана „златният гост“.",
  "Бележка от куратора: Каната е тъмна, защото пази артистична мистерия.",
  "Бележка от куратора: Зелената маса е спокойна, но има мнение за всичко.",
  "Бележка от куратора: Лилавият фон вече си поръча автограф.",
  "Бележка от куратора: Ако картината мълчи, значи мисли много дълбоко.",
  "Бележка от куратора: Крушата поиска червен килим, но получи зелена маса."
];

const curatorButton = document.querySelector("#curatorButton");
const curatorNote = document.querySelector("#curatorNote");
const copyButton = document.querySelector("#copyButton");
const shareText = document.querySelector("#shareText");
const copyStatus = document.querySelector("#copyStatus");
const photoSlots = document.querySelectorAll("[data-photo-src]");

let noteIndex = 0;

curatorButton?.addEventListener("click", () => {
  noteIndex = (noteIndex + 1) % curatorNotes.length;
  curatorNote.textContent = curatorNotes[noteIndex];
  curatorButton.animate(
    [
      { transform: "translateY(0) rotate(0deg)" },
      { transform: "translateY(-5px) rotate(2deg)" },
      { transform: "translateY(0) rotate(0deg)" }
    ],
    {
      duration: 320,
      easing: "ease-out"
    }
  );
});

copyButton?.addEventListener("click", async () => {
  const text = shareText?.textContent?.trim();

  if (!text) {
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    copyStatus.textContent = "Готово! Текстът е копиран.";
  } catch {
    copyStatus.textContent = "Маркирай текста по-горе и го копирай ръчно.";
  }
});

photoSlots.forEach((slot) => {
  const photo = new Image();
  const photoSource = slot.dataset.photoSrc;

  if (!photoSource) {
    return;
  }

  photo.addEventListener("load", () => {
    slot.src = photoSource;
  });

  photo.src = photoSource;
});
