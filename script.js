const processarBtn =
  document.getElementById(
    "processarBtn"
  );

const filtroSelect =
  document.getElementById(
    "filtroSelect"
  );

const dashboard =
  document.getElementById(
    "dashboard"
  );

const tbody =
  document.querySelector(
    "#tabela tbody"
  );

let todasOrdens = [];


// ======================================
// PROCESSAR
// ======================================

processarBtn.addEventListener(
  "click",
  processarArquivo
);

filtroSelect.addEventListener(
  "change",
  aplicarFiltro
);


// ======================================
// PROCESSAR ARQUIVO
// ======================================

function processarArquivo(){

  const fileInput =
    document.getElementById(
      "fileInput"
    );

  if(!fileInput.files.length){

    alert(
      "Selecione uma planilha."
    );

    return;
  }

  const file =
    fileInput.files[0];

  const reader =
    new FileReader();

  reader.onload = function(e){

    const data =
      new Uint8Array(
        e.target.result
      );

    const workbook =
      XLSX.read(data,{
        type:"array"
      });

    const ws =
      workbook.Sheets[
        workbook.SheetNames[0]
      ];

    const json =
      XLSX.utils.sheet_to_json(
        ws,
        {
          header:1,
          defval:""
        }
      );

    processarDados(json);
  };

  reader.readAsArrayBuffer(file);
}


// ======================================
// PROCESSAR DADOS
// ======================================

function processarDados(data){

  todasOrdens = [];

  data.forEach((row,index) => {

    if(index === 0) return;

    const area =
      String(row[0]).trim();

    const tipo =
      String(row[1])
      .toUpperCase()
      .trim();

    const om =
      String(row[3]).trim();

    const texto =
      String(row[4]).trim();

    const centro =
      String(row[5])
      .toUpperCase()
      .trim();

    const statusUsuario =
      String(row[6])
      .toUpperCase();

    const statusSistema =
      String(row[7])
      .toUpperCase();

    // ======================================
    // SOMENTE CAL
    // ======================================

    if(!tipo.includes("CAL")){
      return;
    }

    // ======================================
    // CNPA
    // ======================================

    const cnpa =
      statusSistema.includes(
        "CNPA"
      );

    // ======================================
    // IMPD
    // ======================================

    const impd =
      statusUsuario.includes(
        "IMPD"
      );

    // ======================================
    // CATEGORIA
    // ======================================

    let categoria =
      "OUTROS";

    // TRANSF. e ESTOC.
    if(

      ["CALDTU","MANDTU","AUSETU","RECVALV","OFICOMPL","PINTJCHT"]
      .includes(centro)

      &&

      ["014","018"]
      .includes(area)

    ){

      categoria =
        "TRANSF. e ESTOC.";
    }

    // UTILIDADES
    else if(

      ["CALDTU","MANDTU","AUSETU","RECVALV","OFICOMPL","PINTJCHT"]
      .includes(centro)

      &&

      area === "019"

    ){

      categoria =
        "UTILIDADES";
    }

    // VAZAMENTOS
    else if(

      ["CALDVAZA","MANDVAZA"]
      .includes(centro)

    ){

      categoria =
        "VAZAMENTOS";
    }

    // PURGADORES
    else if(

      ["CALDPURG","MANDPURG"]
      .includes(centro)

    ){

      categoria =
        "PURGADORES";
    }

    // ZR
    else if(

      ["CALDDSC","MANDDSC","AUSEDSC"]
      .includes(centro)

    ){

      categoria =
        "ZR";
    }

    todasOrdens.push({

      area,
      om,
      texto,
      centro,
      categoria,
      cnpa,
      impd
    });
  });

  aplicarFiltro();
}


// ======================================
// FILTRAR
// ======================================

function aplicarFiltro(){

  const filtro =
    filtroSelect.value;

  let filtradas =
    [...todasOrdens];

  // ======================================
  // FILTRO CATEGORIA
  // ======================================

  if(filtro !== "TODOS"){

    filtradas =
      filtradas.filter(
        x =>
          x.categoria === filtro
      );
  }

  // ======================================
  // DASHBOARD
  // ======================================

  atualizarDashboard(
    filtradas,
    filtro
  );

  // ======================================
  // TABELA
  // ======================================

  renderizarTabela(
    filtradas
  );
}


// ======================================
// DASHBOARD
// ======================================

function atualizarDashboard(
  lista,
  filtro
){

  // TOTAL SEM IMPD
  const totalOrdens = lista.length;

  // CNPA SEM IMPD
  const totalCNPA =
    lista.filter(
      x =>
        x.cnpa && !x.impd
    ).length;

  // CNPA + IMPD
  const totalIMPD =
    lista.filter(
      x =>
        x.cnpa && x.impd
    ).length;

  dashboard.innerHTML = `

    <div class="info">

      <h2>
        Área Filtrada
      </h2>

      <p>
        ${filtro}
      </p>

    </div>

    <div class="info">

      <h2>
        Total de Ordens
      </h2>

      <p>
        ${totalOrdens}
      </p>

    </div>

    <div class="info">

      <h2>
        Total CNPA
      </h2>

      <p>
        ${totalCNPA}
      </p>

    </div>

    <div class="info">

      <h2>
        Total CNPA IMPD
      </h2>

      <p>
        ${totalIMPD}
      </p>

    </div>

  `;
}


// ======================================
// TABELA
// ======================================

function renderizarTabela(lista){

  tbody.innerHTML = "";

  lista.forEach(item => {

    const tr =
      document.createElement(
        "tr"
      );

    tr.innerHTML = `

      <td>
        ${item.om}
      </td>

      <td>
        ${item.texto}
      </td>

      <td>
        ${item.area}
      </td>

      <td>
        ${item.centro}
      </td>

      <td>
        ${item.categoria}
      </td>

      <td>

        <span class="
          tag
          ${
            item.cnpa
            ? "cnpa"
            : "nao-cnpa"
          }
        ">

          ${
            item.cnpa
            ? "CNPA"
            : "NÃO"
          }

        </span>

      </td>

      <td>

        <span class="
          tag
          ${
            item.impd
            ? "impd"
            : "normal"
          }
        ">

          ${
            item.impd
            ? "IMPD"
            : "NORMAL"
          }

        </span>

      </td>
    `;

    tbody.appendChild(tr);
  });
}