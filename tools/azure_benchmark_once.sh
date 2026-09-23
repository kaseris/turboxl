#!/usr/bin/env bash
# Provision two Azure benchmark VMs, run the Python comparison, download the
# reports, and delete the entire resource group on every normal/error exit.

set -Eeuo pipefail

readonly SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MODE="csv"
HARNESS_PATH="${SCRIPT_DIR}/cloud_benchmark.py"

LOCATION="uksouth"
ROUNDS=7
ROUNDS_SET=0
SHEET_INDEX=0
OS_TYPE="linux"
RESULTS_ROOT="${PWD}/benchmark-results"
XLSX_PATH="${SCRIPT_DIR}/../benchmarks/turboxl-large-benchmark.xlsx"
XLSX_URL=""
SUBSCRIPTION=""
PANDAS_MANIFEST=""
PANDAS_WHEEL=""
PREPARE_ONLY=0

usage() {
  cat <<'EOF'
Usage:
  tools/azure_benchmark_once.sh [options]

Options:
  --xlsx FILE            Use another local XLSX.
  --xlsx-url URL         Use another XLSX download URL.
  --subscription NAME    Azure subscription name or ID (default: active subscription).
  --os TYPE              Guest OS: linux or windows (default: linux).
  --location REGION      Azure region (default: uksouth).
  --rounds N             Timed rounds per implementation (csv: 7, pandas: 9).
  --sheet-index N        Zero-based worksheet index (default: 0).
  --results-dir DIR      Parent directory for reports (default: ./benchmark-results).
  --mode MODE            csv or pandas (default: csv).
  --manifest FILE        Frozen pandas corpus manifest (pandas mode).
  --wheel FILE           Candidate Linux x86-64 wheel (pandas mode).
  --prepare-only         Validate pandas inputs and package locally; create no Azure resources.
  -h, --help             Show this help.

The script uses the requested subscription, or the active Azure CLI
subscription when none is supplied. It creates two matched 4-vCPU x86-64 VMs
(Intel Standard_D4s_v6 and AMD Standard_D4as_v6), a storage account,
networking, and supporting resources in one temporary resource group. No
inbound ports are opened. Results are downloaded before the group is deleted.
Pandas mode uses Ubuntu 24.04 / Python 3.12; CSV mode keeps Ubuntu 22.04.

The default input is the generated 150,000-row TurboXL benchmark workbook.
EOF
}

while (($#)); do
  case "$1" in
    --xlsx)
      XLSX_PATH="${2:-}"
      XLSX_URL=""
      shift 2
      ;;
    --xlsx-url)
      XLSX_URL="${2:-}"
      XLSX_PATH=""
      shift 2
      ;;
    --location)
      LOCATION="${2:-}"
      shift 2
      ;;
    --subscription)
      SUBSCRIPTION="${2:-}"
      shift 2
      ;;
    --os)
      OS_TYPE="${2:-}"
      shift 2
      ;;
    --rounds)
      ROUNDS="${2:-}"
      ROUNDS_SET=1
      shift 2
      ;;
    --sheet-index)
      SHEET_INDEX="${2:-}"
      shift 2
      ;;
    --results-dir)
      RESULTS_ROOT="${2:-}"
      shift 2
      ;;
    --mode)
      MODE="${2:-}"
      shift 2
      ;;
    --manifest)
      PANDAS_MANIFEST="${2:-}"
      shift 2
      ;;
    --wheel)
      PANDAS_WHEEL="${2:-}"
      shift 2
      ;;
    --prepare-only)
      PREPARE_ONLY=1
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

if [[ "${MODE}" != "csv" && "${MODE}" != "pandas" ]]; then
  echo "--mode must be csv or pandas." >&2
  exit 2
fi
if [[ "${MODE}" == "pandas" ]]; then
  if ((ROUNDS_SET == 0)); then ROUNDS=9; fi
  HARNESS_PATH="${SCRIPT_DIR}/benchmark_pandas.py"
  WHEEL_BLOB_NAME="$(basename "${PANDAS_WHEEL}")"
  if [[ "${OS_TYPE}" != "linux" ]]; then
    echo "Pandas Azure mode currently requires --os linux." >&2
    exit 2
  fi
  if [[ ! -f "${PANDAS_MANIFEST}" || ! -f "${PANDAS_WHEEL}" ||
        "${PANDAS_WHEEL}" != *manylinux*x86_64.whl ]]; then
    echo "Pandas mode requires --manifest and a Linux x86-64 --wheel." >&2
    exit 2
  fi
fi
if ((PREPARE_ONLY)); then
  if [[ "${MODE}" != "pandas" ]]; then
    echo "--prepare-only requires --mode pandas." >&2
    exit 2
  fi
  mkdir -p "${RESULTS_ROOT}"
  python3 "${SCRIPT_DIR}/ci/prepare_pandas_azure.py" \
    --manifest "${PANDAS_MANIFEST}" --wheel "${PANDAS_WHEEL}" \
    --bundle "${RESULTS_ROOT}/pandas-corpus.zip"
  python3 - "${PANDAS_WHEEL}" "${RESULTS_ROOT}/pandas-corpus.zip" <<'PY'
import hashlib
from pathlib import Path
import sys
for filename in sys.argv[1:]:
    path = Path(filename)
    print(f"{path}: {hashlib.sha256(path.read_bytes()).hexdigest()}")
PY
  echo "Pandas Azure inputs prepared; no Azure resources were created."
  exit 0
fi
if [[ "${MODE}" == "csv" && -n "${XLSX_PATH}" && ! -f "${XLSX_PATH}" ]]; then
  echo "--xlsx must point to an existing workbook." >&2
  exit 2
fi
if [[ "${MODE}" == "csv" && -z "${XLSX_PATH}" && -z "${XLSX_URL}" ]]; then
  echo "Provide --xlsx or --xlsx-url." >&2
  exit 2
fi
if [[ ! "${ROUNDS}" =~ ^[1-9][0-9]*$ ]]; then
  echo "--rounds must be a positive integer." >&2
  exit 2
fi
if [[ ! "${SHEET_INDEX}" =~ ^[0-9]+$ ]]; then
  echo "--sheet-index must be a non-negative integer." >&2
  exit 2
fi
if [[ "${OS_TYPE}" != "linux" && "${OS_TYPE}" != "windows" ]]; then
  echo "--os must be linux or windows." >&2
  exit 2
fi
for required_command in az curl python3; do
  if ! command -v "${required_command}" >/dev/null 2>&1; then
    echo "Required command is missing: ${required_command}" >&2
    exit 1
  fi
done
if [[ "${OS_TYPE}" == "linux" ]] && ! command -v ssh-keygen >/dev/null 2>&1; then
  echo "Required command is missing: ssh-keygen" >&2
  exit 1
fi
if [[ ! -f "${HARNESS_PATH}" ]]; then
  echo "Benchmark harness is missing: ${HARNESS_PATH}" >&2
  exit 1
fi

RUN_SUFFIX="$(date -u +%y%m%d%H%M)-$RANDOM"
SAFE_SUFFIX="${RUN_SUFFIX//-/}"
RESOURCE_GROUP="turboxl-bench-${OS_TYPE:0:3}-${RUN_SUFFIX}"
STORAGE_ACCOUNT="txlbench${SAFE_SUFFIX:0:15}"
RESULTS_DIR="${RESULTS_ROOT}/azure-${OS_TYPE}-${RUN_SUFFIX}"
TEMP_DIR="$(mktemp -d "${TMPDIR:-/tmp}/turboxl-azure-bench.XXXXXX")"
RESOURCE_GROUP_CREATED=0
ORIGINAL_SUBSCRIPTION_ID=""
SELECTED_SUBSCRIPTION_ID=""

az_with_retry() {
  local description=$1
  shift
  local attempt
  for attempt in 1 2 3; do
    echo "${description} (attempt ${attempt}/3)..."
    if "$@"; then
      return 0
    fi
    if ((attempt < 3)); then
      echo "Azure returned an error; retrying in 10 seconds..." >&2
      sleep 10
    fi
  done
  echo "Failed: ${description}" >&2
  return 1
}

ensure_provider_registered() {
  local namespace=$1
  local state
  state="$(az provider show \
    --namespace "${namespace}" \
    --subscription "${SUBSCRIPTION_ID}" \
    --query registrationState \
    --output tsv 2>/dev/null || true)"
  if [[ "${state}" == "Registered" ]]; then
    echo "Azure provider ${namespace}: registered"
    return 0
  fi

  echo "Registering Azure provider ${namespace} (current state: ${state:-unknown})..."
  echo "Provider registration is subscription-wide, persists after this run, and has no charge."
  az_with_retry \
    "Registering ${namespace}" \
    az provider register \
      --namespace "${namespace}" \
      --subscription "${SUBSCRIPTION_ID}" \
      --wait \
      --output none
}

cleanup() {
  local exit_code=$?
  trap - EXIT INT TERM
  if ((RESOURCE_GROUP_CREATED)); then
    echo
    echo "Deleting Azure resource group ${RESOURCE_GROUP}..."
    if ! az group delete \
      --name "${RESOURCE_GROUP}" \
      --subscription "${SELECTED_SUBSCRIPTION_ID}" \
      --yes \
      --output none; then
      echo "WARNING: Azure could not confirm deletion. Delete ${RESOURCE_GROUP} in the portal." >&2
      exit_code=1
    else
      echo "Resource group deleted."
    fi
  fi
  if [[ -n "${ORIGINAL_SUBSCRIPTION_ID}" && \
        "${ORIGINAL_SUBSCRIPTION_ID}" != "${SELECTED_SUBSCRIPTION_ID}" ]]; then
    az account set --subscription "${ORIGINAL_SUBSCRIPTION_ID}" >/dev/null 2>&1 || true
  fi
  rm -rf "${TEMP_DIR}"
  exit "${exit_code}"
}
trap cleanup EXIT INT TERM

ORIGINAL_SUBSCRIPTION_ID="$(az account show --query id --output tsv 2>/dev/null || true)"
if [[ -n "${SUBSCRIPTION}" ]]; then
  az account set --subscription "${SUBSCRIPTION}"
fi

SUBSCRIPTION_JSON="$(az account show --output json)" || {
  echo "Azure CLI is not authenticated. Run: az login" >&2
  exit 1
}
SUBSCRIPTION_NAME="$(python3 -c 'import json,sys; print(json.load(sys.stdin)["name"])' <<<"${SUBSCRIPTION_JSON}")"
SUBSCRIPTION_ID="$(python3 -c 'import json,sys; print(json.load(sys.stdin)["id"])' <<<"${SUBSCRIPTION_JSON}")"
SELECTED_SUBSCRIPTION_ID="${SUBSCRIPTION_ID}"

if ! az rest \
  --method get \
  --url "https://management.azure.com/subscriptions/${SUBSCRIPTION_ID}?api-version=2022-12-01" \
  --output none >/dev/null; then
  echo "The current Azure login cannot access subscription ${SUBSCRIPTION_NAME} (${SUBSCRIPTION_ID})." >&2
  echo "Refresh it with: az login --tenant $(python3 -c 'import json,sys; print(json.load(sys.stdin)["tenantId"])' <<<"${SUBSCRIPTION_JSON}")" >&2
  exit 1
fi

echo "Checking required Azure resource providers..."
ensure_provider_registered Microsoft.Storage
ensure_provider_registered Microsoft.Compute
ensure_provider_registered Microsoft.Network
echo

mkdir -p "${RESULTS_DIR}"
if [[ "${MODE}" == "pandas" ]]; then
  python3 "${SCRIPT_DIR}/ci/prepare_pandas_azure.py" \
    --manifest "${PANDAS_MANIFEST}" --wheel "${PANDAS_WHEEL}" \
    --bundle "${TEMP_DIR}/corpus.zip"
  INPUT_DESCRIPTION="${PANDAS_MANIFEST} (validated corpus bundle)"
elif [[ -n "${XLSX_PATH}" ]]; then
  XLSX_PATH="$(cd "$(dirname "${XLSX_PATH}")" && pwd)/$(basename "${XLSX_PATH}")"
  INPUT_DESCRIPTION="${XLSX_PATH}"
else
  XLSX_PATH="${TEMP_DIR}/benchmark.xlsx"
  INPUT_DESCRIPTION="${XLSX_URL}"
  echo "Downloading benchmark workbook before creating Azure resources..."
  curl --fail --location \
    --retry 5 \
    --retry-all-errors \
    --retry-delay 10 \
    --connect-timeout 30 \
    --max-time 900 \
    "${XLSX_URL}" \
    --output "${XLSX_PATH}"
  if ! python3 - "${XLSX_PATH}" <<'PY'
import sys
import zipfile

path = sys.argv[1]
if not zipfile.is_zipfile(path):
    raise SystemExit("Downloaded input is not a valid XLSX/ZIP file.")
with zipfile.ZipFile(path) as workbook:
    if "[Content_Types].xml" not in workbook.namelist():
        raise SystemExit("Downloaded ZIP does not look like an XLSX workbook.")
PY
  then
    exit 1
  fi
fi

echo "TurboXL one-shot Azure benchmark"
echo "Subscription : ${SUBSCRIPTION_NAME} (${SUBSCRIPTION_ID})"
echo "Region       : ${LOCATION}"
echo "Guest OS     : ${OS_TYPE}"
echo "Mode         : ${MODE}"
echo "Resource group: ${RESOURCE_GROUP}"
echo "Input        : ${INPUT_DESCRIPTION}"
echo "Results      : ${RESULTS_DIR}"
if [[ "${OS_TYPE}" == "windows" ]]; then
  echo "Cost note    : Windows Server carries a license premium over the Linux run."
else
  echo "Estimated compute rate: about USD 0.44/hour for both VMs combined."
fi
echo

az_with_retry \
  "Creating resource group ${RESOURCE_GROUP}" \
  az group create \
    --subscription "${SUBSCRIPTION_ID}" \
    --name "${RESOURCE_GROUP}" \
    --location "${LOCATION}" \
    --tags purpose=turboxl-benchmark lifecycle=ephemeral \
    --output none
RESOURCE_GROUP_CREATED=1

az_with_retry \
  "Creating storage account ${STORAGE_ACCOUNT}" \
  az storage account create \
    --subscription "${SUBSCRIPTION_ID}" \
    --name "${STORAGE_ACCOUNT}" \
    --resource-group "${RESOURCE_GROUP}" \
    --location "${LOCATION}" \
    --sku Standard_LRS \
    --kind StorageV2 \
    --allow-blob-public-access false \
    --min-tls-version TLS1_2 \
    --output none

STORAGE_KEY="$(az storage account keys list \
  --resource-group "${RESOURCE_GROUP}" \
  --account-name "${STORAGE_ACCOUNT}" \
  --query '[0].value' --output tsv)"

for container in input results; do
  az storage container create \
    --name "${container}" \
    --account-name "${STORAGE_ACCOUNT}" \
    --account-key "${STORAGE_KEY}" \
    --public-access off \
    --output none
done

if [[ "${MODE}" == "pandas" ]]; then
  for asset in "${TEMP_DIR}/corpus.zip:corpus.zip" "${PANDAS_WHEEL}:${WHEEL_BLOB_NAME}"; do
    az storage blob upload \
      --account-name "${STORAGE_ACCOUNT}" \
      --account-key "${STORAGE_KEY}" \
      --container-name input \
      --name "${asset##*:}" \
      --file "${asset%:*}" \
      --overwrite \
      --output none
  done
else
  az storage blob upload \
    --account-name "${STORAGE_ACCOUNT}" \
    --account-key "${STORAGE_KEY}" \
    --container-name input \
    --name benchmark.xlsx \
    --file "${XLSX_PATH}" \
    --overwrite \
    --output none
fi
az storage blob upload \
  --account-name "${STORAGE_ACCOUNT}" \
  --account-key "${STORAGE_KEY}" \
  --container-name input \
  --name "$(basename "${HARNESS_PATH}")" \
  --file "${HARNESS_PATH}" \
  --overwrite \
  --output none

SAS_EXPIRY="$(python3 -c 'from datetime import datetime,timedelta,timezone; print((datetime.now(timezone.utc)+timedelta(hours=3)).strftime("%Y-%m-%dT%H:%MZ"))')"
INPUT_SAS="$(az storage container generate-sas \
  --name input \
  --account-name "${STORAGE_ACCOUNT}" \
  --account-key "${STORAGE_KEY}" \
  --permissions r \
  --expiry "${SAS_EXPIRY}" \
  --https-only \
  --output tsv)"
OUTPUT_SAS="$(az storage container generate-sas \
  --name results \
  --account-name "${STORAGE_ACCOUNT}" \
  --account-key "${STORAGE_KEY}" \
  --permissions acw \
  --expiry "${SAS_EXPIRY}" \
  --https-only \
  --output tsv)"

if [[ "${MODE}" == "pandas" ]]; then
  DATASET_URL="https://${STORAGE_ACCOUNT}.blob.core.windows.net/input/corpus.zip?${INPUT_SAS}"
  WHEEL_URL="https://${STORAGE_ACCOUNT}.blob.core.windows.net/input/${WHEEL_BLOB_NAME}?${INPUT_SAS}"
else
  DATASET_URL="https://${STORAGE_ACCOUNT}.blob.core.windows.net/input/benchmark.xlsx?${INPUT_SAS}"
fi
HARNESS_URL="https://${STORAGE_ACCOUNT}.blob.core.windows.net/input/$(basename "${HARNESS_PATH}")?${INPUT_SAS}"

WINDOWS_ADMIN_PASSWORD=""
if [[ "${OS_TYPE}" == "linux" ]]; then
  ssh-keygen -q -t ed25519 -N "" -f "${TEMP_DIR}/benchmark_key"
else
  WINDOWS_ADMIN_PASSWORD="$(python3 -c 'import secrets,string; print("Aa1!" + "".join(secrets.choice(string.ascii_letters + string.digits) for _ in range(28)))')"
fi

if [[ "${OS_TYPE}" == "windows" ]]; then
  # Windows computer names are limited to 15 characters.
  declare -a VM_NAMES=("txw-i-${SAFE_SUFFIX:0:6}" "txw-a-${SAFE_SUFFIX:0:6}")
else
  declare -a VM_NAMES=("txl-lin-intel-${SAFE_SUFFIX:0:6}" "txl-lin-amd-${SAFE_SUFFIX:0:6}")
fi
# Matched current-generation 4-vCPU/16-GiB general-purpose VMs. This
# subscription has quota for both v6 families; its equivalent v5 quotas are 0.
declare -a VM_SIZES=("Standard_D4s_v6" "Standard_D4as_v6")
VNET_NAME="txl-benchmark-vnet"
SUBNET_NAME="benchmark"
# This daily schedule is primarily a cost backstop if the local process dies.
AUTO_SHUTDOWN_UTC="$(python3 -c 'from datetime import datetime,timedelta,timezone; print((datetime.now(timezone.utc)+timedelta(hours=2)).strftime("%H%M"))')"

az network vnet create \
  --subscription "${SUBSCRIPTION_ID}" \
  --resource-group "${RESOURCE_GROUP}" \
  --name "${VNET_NAME}" \
  --location "${LOCATION}" \
  --address-prefixes 10.42.0.0/16 \
  --subnet-name "${SUBNET_NAME}" \
  --subnet-prefixes 10.42.1.0/24 \
  --output none

create_vm() {
  local vm_name=$1
  local vm_size=$2
  local -a vm_args=(
    --subscription "${SUBSCRIPTION_ID}" \
    --resource-group "${RESOURCE_GROUP}" \
    --name "${vm_name}" \
    --location "${LOCATION}" \
    --size "${vm_size}" \
    --admin-username azurebench \
    --vnet-name "${VNET_NAME}" \
    --subnet "${SUBNET_NAME}" \
    --public-ip-sku Standard \
    --nsg-rule NONE \
    --storage-sku Standard_LRS \
    --security-type Standard \
    --tags purpose=turboxl-benchmark lifecycle=ephemeral \
    --output none
  )

  if [[ "${OS_TYPE}" == "windows" ]]; then
    vm_args+=(
      --image MicrosoftWindowsServer:WindowsServer:2022-datacenter-azure-edition:latest
      --admin-password "${WINDOWS_ADMIN_PASSWORD}"
      --computer-name "${vm_name}"
      --os-disk-size-gb 127
    )
  else
    local linux_image="Canonical:0001-com-ubuntu-server-jammy:22_04-lts-gen2:latest"
    if [[ "${MODE}" == "pandas" ]]; then
      linux_image="Canonical:ubuntu-24_04-lts:server:latest"
    fi
    vm_args+=(
      --image "${linux_image}"
      --ssh-key-values "${TEMP_DIR}/benchmark_key.pub"
      --os-disk-size-gb 30
    )
  fi

  echo "Validating ${vm_name} (${vm_size}) capacity..."
  az vm create "${vm_args[@]}" --validate
  echo "Creating ${vm_name} (${vm_size})..."
  az vm create "${vm_args[@]}"
  az vm auto-shutdown \
    --subscription "${SUBSCRIPTION_ID}" \
    --resource-group "${RESOURCE_GROUP}" \
    --name "${vm_name}" \
    --time "${AUTO_SHUTDOWN_UTC}" \
    --output none
}

declare -a CREATE_PIDS=()
for index in "${!VM_NAMES[@]}"; do
  create_vm "${VM_NAMES[$index]}" "${VM_SIZES[$index]}" &
  CREATE_PIDS+=("$!")
done
VM_CREATE_FAILED=0
for pid in "${CREATE_PIDS[@]}"; do
  if ! wait "${pid}"; then
    VM_CREATE_FAILED=1
  fi
done
if ((VM_CREATE_FAILED)); then
  echo "At least one VM could not be created in ${LOCATION}." >&2
  echo "Azure capacity changes over time; retry later or choose another region with --location." >&2
  exit 1
fi

run_benchmark() {
  local vm_name=$1
  local vm_size=$2
  local output_url="https://${STORAGE_ACCOUNT}.blob.core.windows.net/results/${vm_name}.json?${OUTPUT_SAS}"
  local dataset_b64 harness_b64 output_b64 remote_script
  dataset_b64="$(printf '%s' "${DATASET_URL}" | base64 | tr -d '\n')"
  harness_b64="$(printf '%s' "${HARNESS_URL}" | base64 | tr -d '\n')"
  output_b64="$(printf '%s' "${output_url}" | base64 | tr -d '\n')"
  if [[ "${OS_TYPE}" == "windows" ]]; then
    remote_script="${TEMP_DIR}/${vm_name}.ps1"
    cat >"${remote_script}" <<EOF
\$ErrorActionPreference = 'Stop'
trap {
  Write-Error ("BENCHMARK_ERROR: " + \$_.Exception.Message)
  Write-Error \$_.ScriptStackTrace
  exit 1
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
\$DatasetUrl = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${dataset_b64}'))
\$HarnessUrl = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${harness_b64}'))
\$OutputUrl = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${output_b64}'))
\$WorkDir = 'C:\turboxl-benchmark'
New-Item -ItemType Directory -Path \$WorkDir -Force | Out-Null

\$PythonInstaller = Join-Path \$WorkDir 'python-3.10.11-amd64.exe'
Invoke-WebRequest -UseBasicParsing \`
  -Uri 'https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe' \`
  -OutFile \$PythonInstaller
\$Install = Start-Process -FilePath \$PythonInstaller -Wait -PassThru \`
  -ArgumentList '/quiet InstallAllUsers=1 PrependPath=0 Include_launcher=0 Include_test=0 Shortcuts=0'
if (\$Install.ExitCode -ne 0) { throw "Python installer exited with code \$(\$Install.ExitCode)" }
\$Python = 'C:\Program Files\Python310\python.exe'
if (-not (Test-Path \$Python)) { throw "Python executable not found at \$Python" }

& \$Python -m pip install --quiet --no-warn-script-location --upgrade pip
if (\$LASTEXITCODE -ne 0) { throw 'pip upgrade failed' }
& \$Python -m pip install --quiet --no-warn-script-location --upgrade turboxl python-calamine openpyxl
if (\$LASTEXITCODE -ne 0) { throw 'benchmark package installation failed' }

\$Workbook = Join-Path \$WorkDir 'benchmark.xlsx'
\$Harness = Join-Path \$WorkDir 'cloud_benchmark.py'
\$ResultPath = Join-Path \$WorkDir 'result.json'
Invoke-WebRequest -UseBasicParsing -Uri \$DatasetUrl -OutFile \$Workbook
Invoke-WebRequest -UseBasicParsing -Uri \$HarnessUrl -OutFile \$Harness

\$env:BENCHMARK_CLOUD_PROVIDER = 'azure'
\$env:BENCHMARK_CLOUD_REGION = '${LOCATION}'
\$env:BENCHMARK_CLOUD_VM_SIZE = '${vm_size}'
\$Result = & \$Python \$Harness \`
  --xlsx \$Workbook \`
  --rounds '${ROUNDS}' \`
  --sheet-index '${SHEET_INDEX}'
if (\$LASTEXITCODE -ne 0) { throw "Benchmark exited with code \$LASTEXITCODE" }
\$Utf8NoBom = New-Object Text.UTF8Encoding(\$false)
[IO.File]::WriteAllText(\$ResultPath, ((\$Result -join [Environment]::NewLine) + [Environment]::NewLine), \$Utf8NoBom)
& \$Python -c "import json; json.load(open(r'C:\turboxl-benchmark\result.json', encoding='utf-8'))"
if (\$LASTEXITCODE -ne 0) { throw 'Result JSON validation failed' }
Invoke-WebRequest -UseBasicParsing \`
  -Method Put \`
  -Headers @{'x-ms-blob-type'='BlockBlob'; 'Content-Type'='application/json'} \`
  -InFile \$ResultPath \`
  -Uri \$OutputUrl
EOF
  else
    remote_script="${TEMP_DIR}/${vm_name}.sh"
    if [[ "${MODE}" == "pandas" ]]; then
      local wheel_b64
      wheel_b64="$(printf '%s' "${WHEEL_URL}" | base64 | tr -d '\n')"
      cat >"${remote_script}" <<EOF
#!/usr/bin/env bash
set -Eeuo pipefail
DATASET_URL="\$(printf '%s' '${dataset_b64}' | base64 -d)"
HARNESS_URL="\$(printf '%s' '${harness_b64}' | base64 -d)"
WHEEL_URL="\$(printf '%s' '${wheel_b64}' | base64 -d)"
OUTPUT_URL="\$(printf '%s' '${output_b64}' | base64 -d)"
sudo apt-get update -qq
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y -qq python3-venv curl
python3 -m venv /tmp/turboxl-benchmark-venv
curl --fail --silent --show-error --location "\${DATASET_URL}" --output /tmp/corpus.zip
curl --fail --silent --show-error --location "\${HARNESS_URL}" --output /tmp/benchmark_pandas.py
curl --fail --silent --show-error --location "\${WHEEL_URL}" --output '/tmp/${WHEEL_BLOB_NAME}'
python3 -m zipfile -e /tmp/corpus.zip /tmp/turboxl-pandas-corpus
/tmp/turboxl-benchmark-venv/bin/python -m pip install --quiet \
  'pandas==3.0.0' 'python-calamine==0.8.2' 'numpy==2.5.3' \
  '/tmp/${WHEEL_BLOB_NAME}'
BENCHMARK_CLOUD_PROVIDER=azure \
BENCHMARK_CLOUD_REGION='${LOCATION}' \
BENCHMARK_CLOUD_VM_SIZE='${vm_size}' \
/tmp/turboxl-benchmark-venv/bin/python /tmp/benchmark_pandas.py \
  --manifest /tmp/turboxl-pandas-corpus/manifest.json \
  --wheel '/tmp/${WHEEL_BLOB_NAME}' \
  --warmups 2 --rounds '${ROUNDS}' --json-output /tmp/result.json || \
  echo 'Benchmark reported a parity or worker failure; uploading its result.' >&2
python3 -c 'import json; json.load(open("/tmp/result.json"))'
curl --fail --silent --show-error \
  --request PUT --header 'x-ms-blob-type: BlockBlob' \
  --header 'Content-Type: application/json' \
  --data-binary @/tmp/result.json "\${OUTPUT_URL}"
EOF
    else
      cat >"${remote_script}" <<EOF
#!/usr/bin/env bash
set -Eeuo pipefail
DATASET_URL="\$(printf '%s' '${dataset_b64}' | base64 -d)"
HARNESS_URL="\$(printf '%s' '${harness_b64}' | base64 -d)"
OUTPUT_URL="\$(printf '%s' '${output_b64}' | base64 -d)"

sudo apt-get update -qq
sudo DEBIAN_FRONTEND=noninteractive apt-get install -y -qq python3-venv curl
python3 -m venv /tmp/turboxl-benchmark-venv
/tmp/turboxl-benchmark-venv/bin/python -m pip install --quiet --upgrade pip
/tmp/turboxl-benchmark-venv/bin/python -m pip install --quiet --upgrade turboxl python-calamine openpyxl
curl --fail --silent --show-error --location "\${DATASET_URL}" --output /tmp/benchmark.xlsx
curl --fail --silent --show-error --location "\${HARNESS_URL}" --output /tmp/cloud_benchmark.py
BENCHMARK_CLOUD_PROVIDER=azure \
BENCHMARK_CLOUD_REGION='${LOCATION}' \
BENCHMARK_CLOUD_VM_SIZE='${vm_size}' \
/tmp/turboxl-benchmark-venv/bin/python /tmp/cloud_benchmark.py \
  --xlsx /tmp/benchmark.xlsx \
  --rounds '${ROUNDS}' \
  --sheet-index '${SHEET_INDEX}' \
  >/tmp/result.json
python3 -c 'import json; json.load(open("/tmp/result.json"))'
curl --fail --silent --show-error \
  --request PUT \
  --header 'x-ms-blob-type: BlockBlob' \
  --header 'Content-Type: application/json' \
  --data-binary @/tmp/result.json \
  "\${OUTPUT_URL}"
EOF
    fi
  fi

  echo "Running benchmark on ${vm_name} (${vm_size})..."
  local attempt
  for attempt in 1 2 3; do
    local command_id="RunShellScript"
    local command_output="${TEMP_DIR}/${vm_name}-run-command-${attempt}.json"
    if [[ "${OS_TYPE}" == "windows" ]]; then
      command_id="RunPowerShellScript"
    fi
    if az vm run-command invoke \
      --subscription "${SUBSCRIPTION_ID}" \
      --resource-group "${RESOURCE_GROUP}" \
      --name "${vm_name}" \
      --command-id "${command_id}" \
      --scripts @"${remote_script}" \
      --output json >"${command_output}"; then
      local blob_exists
      blob_exists="$(az storage blob exists \
        --account-name "${STORAGE_ACCOUNT}" \
        --account-key "${STORAGE_KEY}" \
        --container-name results \
        --name "${vm_name}.json" \
        --query exists \
        --output tsv)"
      if [[ "${blob_exists}" == "true" ]]; then
        echo "Completed ${vm_name}; result upload verified."
        return 0
      fi
      echo "Run Command returned, but ${vm_name} did not upload a result." >&2
      python3 - "${command_output}" <<'PY' >&2
import json
import sys

response = json.load(open(sys.argv[1], encoding="utf-8"))
for entry in response.get("value", []):
    code = entry.get("code", "RunCommand")
    message = entry.get("message", "").strip()
    if message:
        print(f"[{code}]\n{message}")
PY
      echo "Remote benchmark failed on ${vm_name}; not retrying a deterministic script error." >&2
      return 1
    fi
    echo "Run Command attempt ${attempt} failed on ${vm_name}; retrying..." >&2
  done
  echo "Benchmark failed on ${vm_name}." >&2
  return 1
}

declare -a BENCHMARK_PIDS=()
for index in "${!VM_NAMES[@]}"; do
  run_benchmark "${VM_NAMES[$index]}" "${VM_SIZES[$index]}" &
  BENCHMARK_PIDS+=("$!")
done

BENCHMARK_FAILED=0
for pid in "${BENCHMARK_PIDS[@]}"; do
  if ! wait "${pid}"; then
    BENCHMARK_FAILED=1
  fi
done

echo "Downloading available reports..."
az storage blob download-batch \
  --account-name "${STORAGE_ACCOUNT}" \
  --account-key "${STORAGE_KEY}" \
  --source results \
  --destination "${RESULTS_DIR}" \
  --overwrite \
  --output none

if [[ "${MODE}" == "pandas" ]]; then
  python3 - "${RESULTS_DIR}" <<'PY'
import json
from pathlib import Path
import sys

directory = Path(sys.argv[1])
reports = list(sorted(directory.glob("*.json")))
if not reports:
    raise SystemExit("No pandas benchmark reports were downloaded")
for path in reports:
    report = json.loads(path.read_text())
    print(f"{path.name}: real median advantage={report['real_median_advantage']}, "
          f"parity={all(item['parity'] for item in report['workbooks'])}, "
          f"gate={report['passes_20_percent_gate']}")
PY
else
python3 - "${RESULTS_DIR}" <<'PY'
import csv
import json
import pathlib
import sys

directory = pathlib.Path(sys.argv[1])
reports = []
summary_rows = []
for path in sorted(directory.glob("*.json")):
    report = json.loads(path.read_text())
    reports.append(report)
    cpu = report["machine"]["cpu"].get("model_name", "unknown CPU")
    print(f"\n{path.name}: {cpu}")
    for engine, result in report["results"].items():
        print(
            f"  {engine:10} {result['median_seconds']:.3f}s median, "
            f"{result['median_peak_rss_mb']:.1f} MiB peak RSS, "
            f"{result['relative_to_turboxl']:.2f}x TurboXL time"
        )
        summary_rows.append(
            {
                "report": path.name,
                "region": report["machine"]["cloud_region"],
                "vm_size": report["machine"]["cloud_vm_size"],
                "cpu": cpu,
                "engine": engine,
                "version": report["packages"][
                    "python-calamine" if engine == "calamine" else engine
                ],
                "median_seconds": result["median_seconds"],
                "median_peak_rss_mb": result["median_peak_rss_mb"],
                "relative_to_turboxl": result["relative_to_turboxl"],
                "output_sha256": result["sha256"],
            }
        )
    print(f"  output parity: {report['parity']['all_hashes_match']}")
if not reports:
    print("No reports were downloaded.", file=sys.stderr)
    raise SystemExit(1)
with (directory / "summary.csv").open("w", newline="") as output:
    writer = csv.DictWriter(output, fieldnames=summary_rows[0])
    writer.writeheader()
    writer.writerows(summary_rows)
PY
fi

if ((BENCHMARK_FAILED)); then
  echo "At least one VM failed. Any successful report was preserved in ${RESULTS_DIR}." >&2
  exit 1
fi

echo
echo "Reports saved in ${RESULTS_DIR}."
echo "The resource group will now be deleted."
