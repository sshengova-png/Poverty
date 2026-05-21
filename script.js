const curatorNotes = [
  "Curator note: The flowers have requested dramatic lighting and one snack.",
  "Curator note: The vase says its good side is every side.",
  "Curator note: One tulip fainted after hearing the word 'masterpiece'.",
  "Curator note: The imaginary security guard is a butterfly named Kevin.",
  "Curator note: Please clap softly. The roses are practicing confidence.",
  "Curator note: The paintbrush denies all accusations of tickling the canvas."
];

const moodButton = document.querySelector("#moodButton");
const curatorNote = document.querySelector("#curatorNote");

let noteIndex = 0;

moodButton?.addEventListener("click", () => {
  noteIndex = (noteIndex + 1) % curatorNotes.length;
  curatorNote.textContent = curatorNotes[noteIndex];
  moodButton.animate(
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
