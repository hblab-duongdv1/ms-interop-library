# NOTE: This requires Windows containers and a base image with Microsoft Word installed and licensed.
# Microsoft does not support Office automation in server-side unattended scenarios or containers.
# This Dockerfile is illustrative only.

FROM mcr.microsoft.com/dotnet/aspnet:8.0-windowsservercore-ltsc2022 AS base
SHELL ["cmd", "/S", "/C"]

# TODO: Ensure Word is installed in this image or switch to your own custom base image
# that already has Office installed and activated. Example (pseudo):
# FROM yourregistry/office-word-runtime:winservercore-ltsc2022 AS base

WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:8.0-windowsservercore-ltsc2022 AS build
WORKDIR /src

COPY ["DocxToPdf.Api/DocxToPdf.Api.csproj", "DocxToPdf.Api/"]
RUN dotnet restore "DocxToPdf.Api/DocxToPdf.Api.csproj"
COPY . .
WORKDIR "/src/DocxToPdf.Api"
RUN dotnet build "DocxToPdf.Api.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "DocxToPdf.Api.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .

ENV ASPNETCORE_URLS=http://+:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "DocxToPdf.Api.dll"]


