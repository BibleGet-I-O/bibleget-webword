/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
//import "./common-office";
import i18next from 'i18next'
import Backend from 'i18next-http-backend'
import jqueryI18next from 'jquery-i18next'
/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    const lang = Office.context.displayLanguage;ù
    const ns = "bibleget";
    console.log("current Office display language = " + lang);
    i18next.use(Backend).init({
      backend: {
        loadPath: `../../locales/${lang}/${ns}.json`
      }
    }, (err, t) => {
      jqueryI18next.init(i18next, $, {
        tName: 't', // --> appends $.t = i18next.t
        i18nName: 'i18n', // --> appends $.i18n = i18next
        handleName: 'localize', // --> appends $(selector).localize(opts);
        selectorAttr: 'data-i18n', // selector for translating elements
        targetAttr: 'i18n-target', // data-() attribute to grab target element to translate (if different than itself)
        optionsAttr: 'i18n-options', // data-() attribute that contains options, will load/set if useOptionsAttr = true
        useOptionsAttr: false, // see optionsAttr
        parseDefaultValueFromContent: true // parses default values from content ele.val or ele.text
      });

      $('body').localize();

    });
    document.getElementById("getByQuote").onclick = getByQuote;
    document.getElementById("getByKeyword").onclick = getByKeyword;
    document.getElementById("txtQuote").onkeyup = validateQuote;
    document.getElementById("txtFilter").onkeyup = filterSearchResults;
    setPreferedVersionByLang();
  }
});

export async function getByQuote() {
  return Word.run(async context => {
    await getByQuoteOffice(document, context);
    await context.sync();
  });
}

export async function getByKeyword() {
  return Word.run(async context => {
    await getByKeywordOffice(document, context);
    await context.sync();
  });
}

/********************************************** */
export async function getByQuoteOffice(document, context) {
  const quote = document.getElementById("txtQuote").value;
  const version = document.getElementById("cbVersion").value;
  const preferOrigin = getPreferedOrigin();
  console.log("preferOrigin:" + preferOrigin);
  if (quote.trim() != "" && isValidQuote(quote)) {
    await searchByQuote(context.document, quote, version, preferOrigin);
  }
}

function getPreferedOrigin() {
  let rdOrigin = document.getElementsByName("rdOrigin");
  for (let i in rdOrigin) {
    if (rdOrigin[i].checked) {
      return rdOrigin[i].value;
    }
  }
}

function setPreferedVersionByLang() {
  try {
    const lang = "|" + getPreferedLanguage();
    let options = document.querySelectorAll("#cbVersion option");

    for (let i in options) {
      if (typeof options[i].innerHTML != "undefined" && options[i].innerHTML.indexOf(lang) > 0) {
        options[i].selected = true;
        break;
      }
    }
  } catch (e) {
    console.log(e);
  }
}

function getPreferedLanguage() {
  const userLang = navigator.language || navigator.userLanguage;
  if (typeof userLang != "undefined") {
    const idx = userLang.indexOf("-");
    return userLang.substring(idx - 2, idx);
  }
}

async function searchByQuote(document, quote, version, preferOrigin) {
  try {
    const verses = await BibleGetService.getByQuote(quote, version, preferOrigin);
    let range = document.getSelection();

    for (let i in verses) {
      const verse = range.insertText(verses[i].verse + " ", "End");
      verse.font.superscript = true;
      const text = range.insertText(verses[i].text, "End");
      text.font.set({
        name: "Arial",
        bold: false,
        size: 18,
        superscript: false
      });
    }
  } catch (e) {
    notifyError(`Hubo un problema al consultar la cita bíblica en el servidor.`);
    console.error(e);
  }
}

export async function getByKeywordOffice(document, context) {
  const keyword = document.getElementById("txtKeyword").value;
  const exactmatch = document.getElementById("chkExactMatch").checked;
  const version = document.getElementById("cbVersion").value;

  if (keyword.trim() != "") {
    await searchByKeyword(context.document, keyword.trim(), version, exactmatch);
  }
}

var results;
async function searchByKeyword(document, keyword, version, exactmatch) {
  try {
    results = await BibleGetService.searchKeyword(keyword, version, exactmatch);

    if (results.length == 1) {
      insertQuote(document, results[0]);
    } else {
      showResultsSection();
    }
  } catch (e) {
    notifyError(`Hubo un problema al buscar por palabra clave en el servidor.`);
    console.error(e);
  }
}

function insertQuote(document, quote) {
  let range = document.getSelection();
  const verse = range.insertText(quote.verse + " ", "End");
  verse.font.superscript = true;
  const text = range.insertText(quote.text, "End");
  text.font.superscript = false;
}

function showResultsSection() {  
      //Office.context.ui.displayDialogAsync('https://localhost:3000/search-results.html');
      //TODO Show a wait cursor/wheel...
  document.getElementById("appSection").style.display = "none";
  let list = document.getElementById("lstResults");
  for (let i in results) {
    let entry = createResultItem(i);
    list.appendChild(entry);
  }
  document.getElementById("searchResultsSection").style.display = "flex";
}

function createResultItem(i) {  
  const p = document.createElement('p');
  p.classList.add("verse-text");
  const a = document.createElement('a');
  a.classList.add("icon");
  a.classList.add("is-large");
  a.classList.add("ins-result-icon");
  a.onclick = (i) => insertResult(i);
  a.innerHTML = '<i class="fa fa-arrow-circle-down"></i>';
  p.appendChild(a); //TODO Ensayar con tabla... tal vez serìa màs fàcil
  p.appendChild(document.createTextNode(results[i].text));
  const query = document.createElement('span');
  query.classList.add("verse-quote");
  query.appendChild(document.createTextNode(results[i].originalquery));
  p.appendChild(query);
  //a.innerHTML = `<i class="fa fa-arrow-circle-down"></i></a> ${results[i].text} <span class="verse-quote">${results[i].originalquery}</span>`;
  //entry.appendChild(document.createTextNode(res.text));
  //entry.innerHTML = `<p class="verse-text">
  //<a href="#" class="icon is-large ins-result-icon" onclick="insertResult(${i})"><i class="fa fa-arrow-circle-down"></i></a>${results[i].text}
  //<span class="verse-quote">${results[i].originalquery}</span></p>`;
  p.appendChild(a);
  let entry = document.createElement('li');
  entry.appendChild(p);
  return entry;
}

export const insertResult = (i) => {
  console.log(results[i]);
}

/**************************************** */
async function loadVersions() {
  try {
    const versions = await BibleGetService.getVersions();
    let cbVersions = document.getElementById("cbVersion");
    let options = document.querySelectorAll("#cbVersion option");
    options.forEach(o => o.remove());

    for (const key in versions) {
      let opt = document.createElement("option");
      opt.value = key;
      opt.innerHTML = versions[key];
      cbVersions.appendChild(opt);
    }
  } catch (e) {
    notifyError(`Hubo un problema al consultar las versiones en el servidor.`);
    console.error(e);
  }
}
/***************************************************/
export const filterSearchResults = () => {
  const filter = document.getElementById("txtFilter").value.toLowerCase();

  let results = document.querySelectorAll("#lstResults li");
  for (let res of results) {
    if (res.innerHTML.toLowerCase().indexOf(filter) == -1) {
      res.classList.add("hide");
    } else {
      res.classList.remove("hide");
    }
  }
};
/***************************************************/
const regex = /^\d*[a-z]+\d+(([,:]\d+)?(-\d+)?)?$/i;

export const isValidQuote = value => {
  //Bible Verse Regex
  //const regex = /[a-z]+\d+,\d+(-\d+)?/i;
  //const regex = /\d*[a-z]+\d+(?::(\d+))?(-(\d+)(?:([a-z]+)(\d+))?(?::(\d+))?)?/i;
  return regex.test(value.replaceAll(" ", ""));
};
export const validateQuote = () => {
  const value = document.getElementById("txtQuote").value;
  if (value == "" || isValidQuote(value)) {
    document.getElementById("lbErrMsg").innerHTML = ``;
  } else {
    notifyError(`The biblical quote format is not valid.`);
  }
};
function notifyError(errorMessage) {
  document.getElementById("lbErrMsg").innerHTML = errorMessage;
}
/**************************************** */
const axios = require("axios");
//const url = require('url');
const BGET_ENDPOINT = "https://query.bibleget.io/v3/index.php?";
const BGET_SEARCH_ENDPOINT = "https://query.bibleget.io/v3/search.php?";
const BGET_METADATA_ENDPOINT = "https://query.bibleget.io/v3/metadata.php?";

var BibleGetService = {
  getByQuote: async function(quote, version = "CEI2008", preferOrigin = "GREEK") {
    const payload = { query: quote, version: version, preferorigin: preferOrigin, return: "json", appid: "office" };
    //const params = new url.URLSearchParams(payload);
    const response = await axios.get(BGET_ENDPOINT, { params: payload });
    return response.data.results;
  },
  searchKeyword: async function(keyword, version = "CEI2008", exactmatch) {
    const payload = {
      query: "keywordsearch",
      keyword: keyword,
      version: version,
      return: "json",
      appid: "office",
      pluginversion: 1,
      exactmatch: exactmatch
    };
    //const response = await axios.get(BGET_SEARCH_ENDPOINT, { params: payload });
    //return response.data.results;
    return [
      {
        verse: "27",
        text: "Y Dios creó al hombre a su imagen; lo creó a imagen de Dios, los creó varón y mujer.",
        version: "BLPD",
        originalquery: "Gen1:27"
      },
      {
        verse: "2",
        text: "Lorem ipsum doloret sit",
        version: "BLPD",
        originalquery: "Gen1:27"
      },
      {
        verse: "27",
        text: "Beresit bara elohim, we ha ash",
        version: "BLPD",
        originalquery: "Gen1:27"
      }
    ];
  },
  getVersions: async function() {
    const payload = { query: "bibleversions", return: "json" };
    const response = await axios.get(BGET_METADATA_ENDPOINT, { params: payload });
    return response.data.validversions_fullname;
  }
};
