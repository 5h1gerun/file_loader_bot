"""Discord bot that converts Office documents to PDF and PNG images using py-cord.

Requires LibreOffice (soffice) and poppler-utils (pdftoppm) to be installed on the host.
"""

import asyncio
import logging
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List

import discord
from dotenv import load_dotenv
from discord.ext import commands

logging.basicConfig(level=logging.INFO)
LOGGER = logging.getLogger("file-loader-bot")

SUPPORTED_EXTENSIONS = {
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".xlsm",
    ".ppt",
    ".pptx",
    ".odt",
    ".ods",
    ".fods",
    ".odp",
    ".rtf",
    ".pdf",
}

SPREADSHEET_EXTENSIONS = {
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".ods",
    ".fods",
}

OOXML_SPREADSHEET_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

DISCORD_ATTACHMENT_LIMIT = 10


class ConversionError(RuntimeError):
    """Raised when the LibreOffice or pdftoppm conversion fails."""


def _is_supported_attachment(attachment: discord.Attachment) -> bool:
    return Path(attachment.filename).suffix.lower() in SUPPORTED_EXTENSIONS


def _run_subprocess(cmd: list[str], cwd: Path | None = None) -> None:
    LOGGER.debug("Running command: %s", " ".join(cmd))
    completed = subprocess.run(
        cmd,
        cwd=cwd,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if completed.returncode != 0:
        LOGGER.error("Command failed: %s", completed.stderr)
        raise ConversionError(
            f"Command {' '.join(cmd)} failed with exit code {completed.returncode}: "
            f"{completed.stderr.strip()}"
        )
    LOGGER.debug(completed.stdout)


def _ensure_binary(name: str, friendly: str) -> None:
    if shutil.which(name) is None:
        raise ConversionError(
            f"Missing dependency '{name}'. Install {friendly} and restart the bot."
        )


def _prepare_conversion_input(source: Path) -> Path:
    extension = source.suffix.lower()
    if extension in SPREADSHEET_EXTENSIONS:
        return _prepare_spreadsheet_for_single_page(source)
    return source


def _prepare_spreadsheet_for_single_page(source: Path) -> Path:
    prepared_source = source
    if source.suffix.lower() not in OOXML_SPREADSHEET_EXTENSIONS:
        prepared_source = _convert_spreadsheet_to_xlsx(source)
    _fit_workbook_to_single_page(prepared_source)
    return prepared_source


def _convert_spreadsheet_to_xlsx(source: Path) -> Path:
    _ensure_binary("soffice", "LibreOffice")
    target = source.with_suffix(".xlsx")
    _run_subprocess(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(source.parent),
            str(source),
        ],
        cwd=source.parent,
    )
    if not target.exists():
        raise ConversionError(
            f"Spreadsheet conversion to XLSX failed for {source.name}. "
            "LibreOffice did not produce the expected file."
        )
    return target


def _fit_workbook_to_single_page(xlsx_path: Path) -> None:
    try:
        from openpyxl import load_workbook  # type: ignore
        from openpyxl.worksheet.properties import PageSetupProperties  # type: ignore
    except ImportError as exc:  # pragma: no cover - optional dependency
        raise ConversionError(
            "Missing dependency 'openpyxl'. Install it to process spreadsheet attachments."
        ) from exc

    keep_vba = xlsx_path.suffix.lower() == ".xlsm"
    try:
        workbook = load_workbook(xlsx_path, keep_vba=keep_vba)
    except Exception as exc:
        raise ConversionError(f"Failed to load spreadsheet {xlsx_path.name}: {exc}") from exc

    modified = False
    for sheet in workbook.worksheets:
        page_setup = sheet.page_setup
        if (
            page_setup.fitToWidth != 1
            or page_setup.fitToHeight != 1
            or not page_setup.fitToPage
        ):
            page_setup.fitToWidth = 1
            page_setup.fitToHeight = 1
            page_setup.fitToPage = True
            modified = True

        props = sheet.sheet_properties.pageSetUpPr
        if props is None:
            sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            modified = True
        elif props.fitToPage is not True:
            props.fitToPage = True
            modified = True

    if modified:
        workbook.save(xlsx_path)
    workbook.close()


def _convert_to_pdf_sync(input_path: Path) -> Path:
    _ensure_binary("soffice", "LibreOffice")
    output_dir = input_path.parent
    _run_subprocess(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(input_path),
        ]
    )
    pdf_path = output_dir / f"{input_path.stem}.pdf"
    if not pdf_path.exists():
        raise ConversionError("PDF conversion completed but no PDF was produced.")
    return pdf_path


def _pdf_to_images_sync(pdf_path: Path) -> List[Path]:
    _ensure_binary("pdftoppm", "poppler-utils")
    output_base = pdf_path.with_suffix("")
    _run_subprocess(
        ["pdftoppm", "-png", str(pdf_path), str(output_base)],
        cwd=pdf_path.parent,
    )
    images = sorted(pdf_path.parent.glob(f"{pdf_path.stem}-*.png"))
    if not images:
        raise ConversionError("No images were produced from the PDF conversion.")
    return images


async def convert_to_pdf(input_path: Path) -> Path:
    return await asyncio.to_thread(_convert_to_pdf_sync, input_path)


async def pdf_to_images(pdf_path: Path) -> List[Path]:
    return await asyncio.to_thread(_pdf_to_images_sync, pdf_path)


@dataclass
class ConversionArtifacts:
    tmpdir: tempfile.TemporaryDirectory[str]
    workdir: Path
    pdf_path: Path
    image_paths: List[Path]

    def cleanup(self) -> None:
        self.tmpdir.cleanup()


async def build_conversion_artifacts(
    attachment: discord.Attachment,
) -> ConversionArtifacts:
    extension = Path(attachment.filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ConversionError(f"Unsupported file type {extension}")

    tmpdir = tempfile.TemporaryDirectory()
    workdir = Path(tmpdir.name)
    download_path = workdir / attachment.filename
    await attachment.save(download_path)
    conversion_input = await asyncio.to_thread(_prepare_conversion_input, download_path)

    pdf_path = (
        conversion_input
        if conversion_input.suffix.lower() == ".pdf"
        else await convert_to_pdf(conversion_input)
    )
    images = await pdf_to_images(pdf_path)
    return ConversionArtifacts(tmpdir=tmpdir, workdir=workdir, pdf_path=pdf_path, image_paths=images)


intents = discord.Intents.default()
intents.message_content = True
bot = commands.Bot(command_prefix="!", intents=intents)


@bot.event
async def on_ready():
    LOGGER.info("Logged in as %s (ID: %s)", bot.user, bot.user.id)  # type: ignore[attr-defined]


@bot.slash_command(
    name="convert",
    description="Convert an Office document to PDF and PNG previews.",
)
async def convert(
    ctx: discord.ApplicationContext,
    file: discord.Attachment,
):
    extension = Path(file.filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        await ctx.respond(
            f"Unsupported file type `{extension}`. "
            f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}",
            ephemeral=True,
        )
        return

    await ctx.defer()
    artifacts: ConversionArtifacts | None = None
    try:
        artifacts = await build_conversion_artifacts(file)

        pdf_file = discord.File(
            open(artifacts.pdf_path, "rb"), filename=artifacts.pdf_path.name
        )
        try:
            if len(artifacts.image_paths) == 1:
                image_path = artifacts.image_paths[0]
                image_file = discord.File(
                    open(image_path, "rb"),
                    filename=f"{artifacts.pdf_path.stem}_page_1.png",
                )
                try:
                    await ctx.followup.send(
                        content="Here is your PDF and preview image.",
                        files=[pdf_file, image_file],
                    )
                finally:
                    image_file.close()
            else:
                zip_path = artifacts.workdir / f"{artifacts.pdf_path.stem}_images.zip"
                await asyncio.to_thread(_pack_images, artifacts.image_paths, zip_path)
                zip_file = discord.File(open(zip_path, "rb"), filename=zip_path.name)
                try:
                    await ctx.followup.send(
                        content=f"Converted {len(artifacts.image_paths)} pages.",
                        files=[pdf_file, zip_file],
                    )
                finally:
                    zip_file.close()
        finally:
            pdf_file.close()
    except ConversionError as err:
        await ctx.followup.send(f"Conversion failed: {err}", ephemeral=True)
    except Exception:  # pragma: no cover - safeguard
        LOGGER.exception("Unexpected error during conversion")
        await ctx.followup.send("Unexpected error while converting the file.", ephemeral=True)
    finally:
        if artifacts:
            artifacts.cleanup()


@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return

    matching_attachments = [
        attachment for attachment in message.attachments if _is_supported_attachment(attachment)
    ]
    for attachment in matching_attachments:
        await _handle_auto_conversion(message, attachment)

    await bot.process_commands(message)


async def _handle_auto_conversion(
    message: discord.Message,
    attachment: discord.Attachment,
) -> None:
    artifacts: ConversionArtifacts | None = None
    try:
        async with message.channel.typing():
            artifacts = await build_conversion_artifacts(attachment)
        await _send_image_batches(message, attachment.filename, artifacts.image_paths)
    except ConversionError as err:
        await message.reply(
            f"Failed to convert `{attachment.filename}`: {err}",
            mention_author=False,
        )
    except Exception:  # pragma: no cover - safeguard
        LOGGER.exception("Unexpected error while auto-converting attachment")
        await message.reply(
            f"Unexpected error while converting `{attachment.filename}`.",
            mention_author=False,
        )
    finally:
        if artifacts:
            artifacts.cleanup()


async def _send_image_batches(
    message: discord.Message,
    original_name: str,
    image_paths: List[Path],
) -> None:
    if not image_paths:
        raise ConversionError("No pages were produced during conversion.")

    pending: list[tuple[discord.File, int]] = []
    first_chunk = True
    for page_number, image_path in enumerate(image_paths, start=1):
        file_handle = discord.File(
            open(image_path, "rb"),
            filename=f"{Path(original_name).stem}_page_{page_number}.png",
        )
        pending.append((file_handle, page_number))
        if len(pending) == DISCORD_ATTACHMENT_LIMIT:
            await _dispatch_chunk(message, original_name, pending, first_chunk)
            first_chunk = False
            pending = []

    if pending:
        await _dispatch_chunk(message, original_name, pending, first_chunk)


async def _dispatch_chunk(
    message: discord.Message,
    original_name: str,
    chunk: list[tuple[discord.File, int]],
    first_chunk: bool,
) -> None:
    start_page = chunk[0][1]
    end_page = chunk[-1][1]
    page_desc = (
        f"page {start_page}" if start_page == end_page else f"pages {start_page}-{end_page}"
    )
    content = f"`{original_name}` {page_desc}"
    files = [file for file, _ in chunk]
    try:
        if first_chunk:
            await message.reply(content=content, files=files, mention_author=False)
        else:
            await message.channel.send(content=content, files=files)
    finally:
        for file, _ in chunk:
            file.close()


def _pack_images(images: List[Path], zip_path: Path) -> None:
    import zipfile

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for image in images:
            archive.write(image, arcname=image.name)


def main() -> None:
    load_dotenv()
    token = os.getenv("DISCORD_BOT_TOKEN")
    if not token:
        raise RuntimeError("Set the DISCORD_BOT_TOKEN environment variable.")
    bot.run(token)


if __name__ == "__main__":
    main()
