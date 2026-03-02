import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Packer, Paragraph, Document, ImageRun } from 'docx';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';

(pdfjsLib as any).GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.5.136/pdf.worker.min.mjs';

interface Aluno {
  nome: string;
  rm: string;
  usuario: string;
  email: string;
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  unidade = '/Alunos/Alunos - centro';
  dominio = 'seudominio.com.br';
  turma = 'manual';
  textoNomes = '';
  alunos: Aluno[] = [];
  status = '';
  imagemInfo?: Uint8Array;

  readonly unidadesDisponiveis = [
    '/Alunos/Alunos - centro',
    '/Alunos/Alunos - cohab',
    '/Alunos/Alunos - raposa',
    '/Unidade Bacabal/Alunos - bacabal'
  ];

  async onPdfSelecionado(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.turma = file.name.replace(/\.pdf$/i, '').trim() || 'turma';
    const bytes = new Uint8Array(await file.arrayBuffer());

    const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
    const linhas: string[] = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const pageText = textContent.items
        .map((item: any) => item.str ?? '')
        .join('\n');
      linhas.push(...pageText.split(/\r?\n/));
    }

    const alunosBase = this.extrairAlunos(linhas);
    this.alunos = this.atribuirCredenciais(alunosBase);
    this.status = `${this.alunos.length} alunos extraídos do PDF.`;
  }

  processarTexto(): void {
    this.turma = 'manual';
    const linhas = this.textoNomes.split(/\r?\n/).map((linha) => linha.trim()).filter(Boolean);
    const alunosBase = linhas.map((nome) => ({ nome, rm: '' }));
    this.alunos = this.atribuirCredenciais(alunosBase);
    this.status = `${this.alunos.length} alunos preparados pelo modo manual.`;
  }

  async onImagemSelecionada(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) {
      this.imagemInfo = undefined;
      return;
    }
    this.imagemInfo = new Uint8Array(await file.arrayBuffer());
  }

  baixarCsv(): void {
    if (!this.alunos.length) {
      this.status = 'Nenhum aluno para exportar.';
      return;
    }

    const cabecalho = 'Nome;Usuario;Email;Turma;Unidade';
    const linhas = this.alunos.map((a) => `${a.nome};${a.usuario};${a.email};${this.turma};${this.unidade}`);
    const conteudo = [cabecalho, ...linhas].join('\n');

    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8' });
    saveAs(blob, `usuarios_${this.turma}.csv`);
    this.status = 'CSV gerado com sucesso.';
  }

  async baixarDocx(): Promise<void> {
    if (!this.alunos.length) {
      this.status = 'Nenhum aluno para exportar.';
      return;
    }

    const sections = this.alunos.map((aluno) => {
      const children: Paragraph[] = [
        new Paragraph(`Aluno: ${aluno.nome}`),
        new Paragraph(`RM: ${aluno.rm || '-'}`),
        new Paragraph(`Usuário: ${aluno.usuario}`),
        new Paragraph('Senha padrão: 123456'),
        new Paragraph('')
      ];

      if (this.imagemInfo) {
        children.push(
          new Paragraph({
            children: [
              new ImageRun({
                data: this.imagemInfo,
                transformation: { width: 180, height: 180 }
              })
            ]
          })
        );
      }

      return { children };
    });

    const doc = new Document({ sections });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `usuarios_${this.turma}.docx`);
    this.status = 'DOCX gerado com sucesso.';
  }

  private extrairAlunos(linhasOriginais: string[]): Array<{ nome: string; rm: string }> {
    const linhas = linhasOriginais.map((l) => l.trim()).filter(Boolean);
    const alunos: Array<{ nome: string; rm: string }> = [];

    for (let i = 0; i < linhas.length; i++) {
      const linha = linhas[i];
      const matchNome = linha.match(/^([A-ZÁÉÍÓÚÃÕÂÊÎÔÛÇ ]{5,})(?:\s+(\d{6,8}))?$/);
      if (!matchNome) continue;

      const nome = matchNome[1].replace(/\s+/g, ' ').trim();
      let rm = matchNome[2] ?? '';

      if (!rm && i + 1 < linhas.length) {
        const rmMatch = linhas[i + 1].match(/\b(\d{6,8})\b/);
        if (rmMatch) {
          rm = rmMatch[1];
          i += 1;
        }
      }

      alunos.push({ nome, rm });
    }

    return alunos;
  }

  private atribuirCredenciais(alunosBase: Array<{ nome: string; rm: string }>): Aluno[] {
    const usados = new Set<string>();

    return alunosBase.map((aluno) => {
      const base = this.normalizarUsuario(aluno.nome);
      let usuario = base;
      let sufixo = 1;

      while (usados.has(usuario)) {
        sufixo += 1;
        usuario = `${base}${sufixo}`;
      }
      usados.add(usuario);

      return {
        ...aluno,
        usuario,
        email: `${usuario}@${this.dominio}`
      };
    });
  }

  private normalizarUsuario(nome: string): string {
    const semAcento = nome
      .normalize('NFKD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z\s]/g, '')
      .trim();

    const partes = semAcento.split(/\s+/).filter(Boolean).map((p) => p.toLowerCase());
    if (!partes.length) return 'usuario';
    if (partes.length === 1) return partes[0];
    return `${partes[0]}.${partes[partes.length - 1]}`;
  }
}
