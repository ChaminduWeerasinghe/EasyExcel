FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
WORKDIR /src
COPY ["EasyExcelTest/EasyExcelTest.csproj", "EasyExcelTest/"]
RUN dotnet restore "EasyExcelTest/EasyExcelTest.csproj"
COPY . .
WORKDIR "/src/EasyExcelTest"
RUN dotnet build "EasyExcelTest.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "EasyExcelTest.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "EasyExcelTest.dll"]
