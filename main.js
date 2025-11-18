let value = 0;
let direction = 1;

setInterval(() => {
  value += direction;

  if (value >= 60 || value <= 0) {
    direction *= -1;
  }

  document.querySelectorAll(".morph").forEach((el) => {
    el.style.fontVariationSettings = `"MORF" ${value}`;
  });
}, 10);

const phrases = [
  "“How do I add page numbers?”",
  "“Where is the line spacing option?”",
  "“How can I insert a header?”",
  "“How do I change page orientation?”",
  "“Where can I find the Styles panel?”",
  "“How to insert a page break?”",
  "“How do I add bullets?”",
  "“How to adjust margins?”",
  "“Where is the Track Changes tool?”",
  "“How to insert a picture?”",
  "“Why can’t I select the text?”",
  "“How do I wrap text around an image?”",
  "“Where is the spell check?”",
  "“How to insert a hyperlink?”",
  "“Why is my ruler not showing?”",
  "“How do I create a table of contents?”",
  "“How to change font size quickly?”",
  "“Where is the Format Painter?”",
  "“How do I remove extra spaces?”",
  "“How to insert shapes?”",
  "“Why are my paragraphs uneven?”",
  "“How do I split a table?”",
  "“How to convert text to a table?”",
  "“Where is the Word Count tool?”",
  "“How do I save as PDF?”",
  "“How to add comments?”",
  "“Why is my text centered?”",
  "“How do I enable dark mode?”",
  "“How to insert symbols?”",
  "“Where is the watermark option?”",
];

let index = 0;
let charIndex = 0;
let element = document.getElementById("typer");

function type() {
  let current = phrases[index];

  element.textContent = current.substring(0, charIndex);
  charIndex++;

  if (charIndex <= current.length) {
    setTimeout(type, 80);
  } else {
    setTimeout(erase, 1500);
  }
}

function erase() {
  let current = phrases[index];

  element.textContent = current.substring(0, charIndex);
  charIndex--;

  if (charIndex >= 0) {
    setTimeout(erase, 40);
  } else {
    index = (index + 1) % phrases.length;
    setTimeout(type, 300);
  }
}

type();

function toggleMenu() {
  const menu = document.querySelector(".menuBox");
  const icon = document.querySelector(".menuIconP");

  if (!menu.classList.contains("show")) {
    menu.classList.add("show");

    requestAnimationFrame(() => {
      menu.classList.add("moved");
    });
  } else {
    menu.classList.remove("moved");

    menu.addEventListener(
      "transitionend",
      () => {
        menu.classList.remove("show");
      },
      { once: true }
    );
  }

  icon.classList.toggle("rotated");
}
