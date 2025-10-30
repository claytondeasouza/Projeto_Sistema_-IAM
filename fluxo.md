graph TD
    subgraph "Início do Processo"
        A[Formulário de Solicitação de Acesso preenchido] --> B{Arquivo Excel gerado};
        B --> C[Salvo na pasta 'Formulários IAM' no SharePoint];
    end

    subgraph "Power Automate: Orquestração"
        C -- Gatilho: Novo arquivo criado --> D[1. Ler linhas do arquivo Excel];
        D --> E[2. Iniciar e declarar variáveis<br>(Nome, Cargo, Gestor, etc.)];
        E --> F[3. Criar Chamado no Qualitor via API REST<br>Anexar Excel];
        F --> G[4. Enviar para Aprovação do Gestor<br>Via Teams e Outlook];
    end

    subgraph "Decisão do Gestor"
        G --> H{Gestor Aprovou?};
    end

    subgraph "Trilha de Rejeição"
        H -- ❌ Não --> I[a. Atualizar chamado no Qualitor<br>Status: Rejeitado];
        I --> J[b. Notificar Solicitante<br>Via E-mail/Teams];
        J --> K[Fim do Fluxo];
    end

    subgraph "Trilha de Aprovação"
        H -- ✅ Sim --> L[5. Executar Script PowerShell via Gateway Local];
        L --> M{Tipo de Solicitação?};
        M -- Contratação --> N[a. Script: Criar usuário no AD<br>Adicionar a grupos];
        M -- Alteração --> O[b. Script: Modificar usuário no AD<br>Ajustar grupos];
        M -- Desligamento --> P[c. Script: Desabilitar usuário no AD<br>Remover de grupos];

        N --> Q[6. Atribuir Licença do Office 365<br>Via Microsoft Graph API];
        O --> Q;
        P --> Q;

        Q --> R[7. Registrar Log de Auditoria<br>Gravar no Banco de Dados SQL];
        R --> S[8. Atualizar Chamado no Qualitor<br>Status: Finalizado com Sucesso];
        S --> T[Fim do Fluxo];
    end

    style F fill:#0078D4,color:#fff
    style G fill:#0078D4,color:#fff
    style L fill:#2E8B57,color:#fff
    style Q fill:#2E8B57,color:#fff
    style R fill:#DAA520,color:#fff
    style S fill:#0078D4,color:#fff
    style I fill:#DC143C,color:#fff