"""Downloads and extract the Visual C++ Redistributables.

This is useful when working with minidump files as the user may be running
with a new different version of the runtime. This script aims to maintain
a copy of the various versions.

Versions are normally added once I encounter them, in November 2022, I added
a rather large back catalogue of versions.

This requires dark.exe from Wix (http://wixtoolset.org/releases/) and
expand.exe (File Expansion Utility - this comes with Microsoft Windows).

For extracting version 9 (Visual C++ 2008) it requires 7zr and 7za from 7-zip.

Known Limitations:
- No extraction method for Visual C++ 6, 8, and 10 are implemented.
  Version 6 and 8 likely would require the ability to extract a MSI.
- Visual C++ 9, 11 and 12 don't extract with nicely named files. For example,
  a file is named F_CENTRAL_msvcr120_x64 rather than msvcr120.dll.
"""
from __future__ import print_function

import argparse
import os
import requests
import subprocess
import tempfile
import zipfile


def download_file(uri, dest):
    """Download the uri to dest."""
    if os.path.isfile(dest):
        # Skipping, already downloaded.
        return dest
    r = requests.get(uri, stream=True)
    with open(dest, 'wb') as f:
        for chunk in r.iter_content(chunk_size=1024):
            if chunk: # filter out keep-alive new chunks
                f.write(chunk)
    return dest

def fetch_runtimes(download_directory, include_old_versions):
    """Fetches the Visual C++ x64 Redistributable from Microsoft.

    download_directory: The directory to write the resulting files in.
    include_old_versions: If true versions older than 14 will be fetched.

    Return the (version, runtime_file) for each known runtime.
    """
    # Some of these packages came from:
    # https://blogs.technet.microsoft.com/jagbal/2017/09/04/where-can-i-download-the-old-visual-c-redistributables/
    # Others came from Chocolatey (https://chocolatey.org/)
    runtime_downloads = [
        (
            '14.42.34438.0',
            'https://download.visualstudio.microsoft.com/download/pr/285b28c7-3cf9-47fb-9be8-01cf5323a8df/8F9FB1B3CFE6E5092CF1225ECD6659DAB7CE50B8BF935CB79BFEDE1F3C895240/VC_redist.x64.exe'
        )
        # (
        #     '6.0.2900.2180',  # Visual C++ 6.
        #     'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE',
        # ),
        # (
        #     '8.0.50727.619501',  # Visual C++ 2005.
        #     'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE',
        # ),
        # (
        #     '9.0.30729.4148',  # Visual C++ 2008.
        #     'https://download.microsoft.com/download/9/7/7/977B481A-7BA6-4E30-AC40-ED51EB2028F2/vcredist_x64.exe',
        # ),
        # (
        #     '10.0.40219.32503', # Visual C++ 2010.
        #     'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe',
        # ),
        # (
        #     '11.0.61030.0', # Visual C++ 2012.
        #     'http://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe',
        # ),
        # (
        #     '12.0.30501.0', # Visual C++ 2013.
        #     'http://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe',
        # ),
        # # Visual C++ 2015 to 2022.
        # (
        #     '14.0.23026.0',
        #     'https://download.microsoft.com/download/9/3/F/93FCF1E7-E6A4-478B-96E7-D4B285925B00/vc_redist.x64.exe',
        #     # This next one seems to be the same version.
        #     #'https://download.microsoft.com/download/4/7/a/47a9d485-9e3a-4707-be1d-26011ff15f42/vc_redist.x64.exe',
        # ),
        # (
        #     '14.0.25325.0',
        #     'https://download.visualstudio.microsoft.com/download/pr/11100230/15ccb3f02745c7b206ad10373cbca89b/VC_redist.x64.exe',
        # ),
        # (
        #     '14.10.25017',
        #     'https://download.microsoft.com/download/3/b/f/3bf6e759-c555-4595-8973-86b7b4312927/vc_redist.x64.exe',
        # ),
        # (
        #     '14.10.25008.0',
        #     'https://download.microsoft.com/download/8/9/d/89d195e1-1901-4036-9a75-fbe46443fc5a/vc_redist.x64.exe',
        # ),
        # (
        #     '14.10.25017.0',
        #     'https://download.microsoft.com/download/3/b/f/3bf6e759-c555-4595-8973-86b7b4312927/vc_redist.x64.exe',
        # ),
        # (
        #     '14.11.25325.0',
        #     'https://download.visualstudio.microsoft.com/download/pr/11100230/15ccb3f02745c7b206ad10373cbca89b/VC_redist.x64.exe',
        # ),
        # (
        #     '14.12.25810',
        #     'https://download.visualstudio.microsoft.com/download/pr/100349091/2cd2dba5748dc95950a5c42c2d2d78e4/VC_redist.x64.exe',
        # ),
        # (
        #     '14.12.2581',
        #     'https://download.visualstudio.microsoft.com/download/pr/100349091/2cd2dba5748dc95950a5c42c2d2d78e4/VC_redist.x64.exe',
        # ),
        # (
        #     '14.15.26706',
        #     'https://download.visualstudio.microsoft.com/download/pr/20ef12bb-5283-41d7-90f7-eb3bb7355de7/8b58fd89f948b2430811db3da92299a6/vc_redist.x64.exe',
        # ),
        # (
        #     '14.16.27012',
        #     'https://download.visualstudio.microsoft.com/download/pr/9fbed7c7-7012-4cc0-a0a3-a541f51981b5/e7eec15278b4473e26d7e32cef53a34c/vc_redist.x64.exe',
        # ),
        # (
        #     '14.16.27024',
        #     'https://download.visualstudio.microsoft.com/download/pr/46db022e-06ea-4d11-a724-d26d33bc63f7/2b635c854654745078d5577a8ed0f80d/vc_redist.x64.exe',
        # ),
        # (
        #     '14.16.27027',
        #     'https://download.visualstudio.microsoft.com/download/pr/36c5faaf-bd8b-433f-b3d7-2af73bae10a8/212f41f2ccffee6d6dc27f901b7d77a1/vc_redist.x64.exe',
        # ),
        # (
        #     '14.16.27033',
        #     'https://download.visualstudio.microsoft.com/download/pr/4100b84d-1b4d-487d-9f89-1354a7138c8f/5B0CBB977F2F5253B1EBE5C9D30EDBDA35DBD68FB70DE7AF5FAAC6423DB575B5/VC_redist.x64.exe',
        # ),
        # (
        #     '14.20.27508',
        #     'https://download.visualstudio.microsoft.com/download/pr/21614507-28c5-47e3-973f-85e7f66545a4/f3a2caa13afd59dd0e57ea374dbe8855/vc_redist.x64.exe',
        # ),
        # (
        #     '14.21.27702',
        #     'https://download.visualstudio.microsoft.com/download/pr/9e04d214-5a9d-4515-9960-3d71398d98c3/1e1e62ab57bbb4bf5199e8ce88f040be/vc_redist.x64.exe',
        # ),
        # (
        #     '14.22.27821',
        #     'https://download.visualstudio.microsoft.com/download/pr/cc0046d4-e7b4-45a1-bd46-b1c079191224/9c4042a4c2e6d1f661f4c58cf4d129e9/vc_redist.x64.exe',
        # ),
        # (
        #     '14.23.27820',
        #     'https://download.visualstudio.microsoft.com/download/pr/9565895b-35a6-434b-a881-11a6f4beec76/EE84FED2552E018E854D4CD2496DF4DD516F30733A27901167B8A9882119E57C/VC_redist.x64.exe',
        # ),
        # (
        #     '14.24.28127',
        #     'https://download.visualstudio.microsoft.com/download/pr/3b070396-b7fb-4eee-aa8b-102a23c3e4f4/40EA2955391C9EAE3E35619C4C24B5AAF3D17AEAA6D09424EE9672AA9372AEED/VC_redist.x64.exe',
        # ),
        # (
        #     '14.25.28508',
        #     'https://download.visualstudio.microsoft.com/download/pr/8c211be1-c537-4402-82e7-a8fb5ee05e8a/B6C82087A2C443DB859FDBEAAE7F46244D06C3F2A7F71C35E50358066253DE52/VC_redist.x64.exe',
        # ),
        # (
        #     '14.26.28720',
        #     'https://download.visualstudio.microsoft.com/download/pr/d60aa805-26e9-47df-b4e3-cd6fcc392333/7D7105C52FCD6766BEEE1AE162AA81E278686122C1E44890712326634D0B055E/VC_redist.x64.exe',
        # ),
        # (
        #     '14.27.29112',
        #     'https://download.visualstudio.microsoft.com/download/pr/48431a06-59c5-4b63-a102-20b66a521863/4B5890EB1AEFDF8DFA3234B5032147EB90F050C5758A80901B201AE969780107/VC_redist.x64.exe',
        # ),
        # (
        #     '14.28.29325',
        #     'https://download.visualstudio.microsoft.com/download/pr/89a3b9df-4a09-492e-8474-8f92c115c51d/B1A32C71A6B7D5978904FB223763263EA5A7EB23B2C44A0D60E90D234AD99178/VC_redist.x64.exe',
        # ),
        # (
        #     '14.28.29910',
        #     'https://download.visualstudio.microsoft.com/download/pr/cd3a705f-70b6-46f7-b8e2-63e6acc5bd05/F299953673DE262FEFAD9DD19BFBE6A5725A03AE733BEBFEC856F1306F79C9F7/VC_redist.x64.exe',
        # ),
        # (
        #     '14.28.29913',
        #     'https://download.visualstudio.microsoft.com/download/pr/366c0fb9-fe05-4b58-949a-5bc36e50e370/015EDD4E5D36E053B23A01ADB77A2B12444D3FB6ECCEFE23E3A8CD6388616A16/VC_redist.x64.exe',
        # ),
        # (
        #     '14.28.29914',
        #     'https://download.visualstudio.microsoft.com/download/pr/85d47aa9-69ae-4162-8300-e6b7e4bf3cf3/52B196BBE9016488C735E7B41805B651261FFA5D7AA86EB6A1D0095BE83687B2/VC_redist.x64.exe',
        # ),
        # (
        #     '14.29.30037',
        #     'https://download.visualstudio.microsoft.com/download/pr/f1998402-3cc0-466f-bd67-d9fb6cd2379b/A1592D3DA2B27230C087A3B069409C1E82C2664B0D4C3B511701624702B2E2A3/VC_redist.x64.exe',
        # ),
        # (
        #     '14.29.30040',
        #     'https://download.visualstudio.microsoft.com/download/pr/36e45907-8554-4390-ba70-9f6306924167/97CC5066EB3C7246CF89B735AE0F5A5304A7EE33DC087D65D9DFF3A1A73FE803/VC_redist.x64.exe',
        # ),
        # (
        #     '14.29.30133',
        #     'https://download.visualstudio.microsoft.com/download/pr/7239cdc3-bd73-4f27-9943-22de059a6267/003063723B2131DA23F40E2063FB79867BAE275F7B5C099DBD1792E25845872B/VC_redist.x64.exe',
        # ),
        # (
        #     '14.29.30135',
        #     'https://download.visualstudio.microsoft.com/download/pr/d3cbdace-2bb8-4dc5-a326-2c1c0f1ad5ae/9B9DD72C27AB1DB081DE56BB7B73BEE9A00F60D14ED8E6FDE45DAB3E619B5F04/VC_redist.x64.exe',
        # ),
        # (
        #     '14.29.30139',
        #     'https://download.visualstudio.microsoft.com/download/pr/b929b7fe-5c89-4553-9abe-6324631dcc3a/296F96CD102250636BCD23AB6E6CF70935337B1BBB3507FE8521D8D9CFAA932F/VC_redist.x64.exe',
        # ),
        # (
        #     '14.30.30704',
        #     'https://download.visualstudio.microsoft.com/download/pr/10a8d53a-c69e-4586-8c6b-c416bf85a0ae/A9F5D2EAF67BF0DB0178B6552A71C523C707DF0E2CC66C06BFBC08BDC53387E7/VC_redist.x64.exe',
        # ),
        # (
        #     '14.30.30708',
        #     'https://download.visualstudio.microsoft.com/download/pr/571ad766-28d1-4028-9063-0fa32401e78f/5D3D8C6779750F92F3726C70E92F0F8BF92D3AE2ABD43BA28C6306466DE8A144/VC_redist.x64.exe',
        # ),
        # (
        #     '14.31.31103',
        #     'https://download.visualstudio.microsoft.com/download/pr/d22ecb93-6eab-4ce1-89f3-97a816c55f04/37ED59A66699C0E5A7EBEEF7352D7C1C2ED5EDE7212950A1B0A8EE289AF4A95B/VC_redist.x64.exe',
        # ),
        # (
        #     '14.32.31326',
        #     'https://download.visualstudio.microsoft.com/download/pr/6b6923b0-3045-4379-a96f-ef5506a65d5b/426A34C6F10EA8F7DA58A8C976B586AD84DD4BAB42A0CFDBE941F1763B7755E5/VC_redist.x64.exe',
        # ),
        # (
        #     '14.32.31332',
        #     'https://download.visualstudio.microsoft.com/download/pr/ed95ef9e-da02-4735-9064-bd1f7f69b6ed/CE6593A1520591E7DEA2B93FD03116E3FC3B3821A0525322B0A430FAA6B3C0B4/VC_redist.x64.exe',
        # ),
        # (
        #     '14.34.31931',
        #     'https://download.visualstudio.microsoft.com/download/pr/bcb0cef1-f8cb-4311-8a5c-650a5b694eab/2257B3FBE3C7559DE8B31170155A433FAF5B83829E67C589D5674FF086B868B9/VC_redist.x64.exe',
        # ),
        # (
        #     '14.34.31938',
        #     'https://download.visualstudio.microsoft.com/download/pr/8b92f460-7e03-4c75-a139-e264a770758d/26C2C72FBA6438F5E29AF8EBC4826A1E424581B3C446F8C735361F1DB7BEFF72/VC_redist.x64.exe',
        # ),
        # (
        #     '14.36.32532',
        #     'https://download.visualstudio.microsoft.com/download/pr/eaab1f82-787d-4fd7-8c73-f782341a0c63/917C37D816488545B70AFFD77D6E486E4DD27E2ECE63F6BBAAF486B178B2B888/VC_redist.x64.exe',
        # ),
        # (
        #     '14.38.33130',
        #     'https://download.visualstudio.microsoft.com/download/pr/a061be25-c14a-489a-8c7c-bb72adfb3cab/4DFE83C91124CD542F4222FE2C396CABEAC617BB6F59BDCBDF89FD6F0DF0A32F/VC_redist.x64.exe',
        # ),
        # (
        #     '14.38.33135',
        #     'https://download.visualstudio.microsoft.com/download/pr/6ba404bb-6312-403e-83be-04b062914c98/1AD7988C17663CC742B01BEF1A6DF2ED1741173009579AD50A94434E54F56073/VC_redist.x64.exe'
        # ),
        # (
        #     '14.40.33810',
        #     'https://download.visualstudio.microsoft.com/download/pr/1754ea58-11a6-44ab-a262-696e194ce543/3642E3F95D50CC193E4B5A0B0FFBF7FE2C08801517758B4C8AEB7105A091208A/VC_redist.x64.exe'
        # ),
    ]

    for version, download_uri in runtime_downloads:
        if not include_old_versions and int(version.partition('.')[0]) < 14:
            # Don't include versions older than version 14.
            continue

        filename = version + '_' + os.path.basename(download_uri)
        destination = download_file(download_uri, os.path.join(download_directory, filename))
        yield version, destination


def fetch_7zip(download_directory):
    """Downloads the 7-zip console executable.

    This is used to extract versions of runtime packages prior to the use of
    WiX Burn.
    """
    # First, download the console executable which can extract a 7z file.
    # Next, it downloads a 7z and extracts that which contains a more fully
    # fledged console executable that can extract the CAB files in the old
    # executables.
    exe_path = os.path.join(download_directory, '__7zr.exe')
    if not os.path.isfile(exe_path):
        download_file('https://7-zip.org/a/7zr.exe', exe_path)

    tool_archive = os.path.join(download_directory, '7z2301-extra.7z')
    if not os.path.isfile(tool_archive):
        download_file('https://7-zip.org/a/7z2301-extra.7z', tool_archive)

    # Extract the 7z file.
    tools_path = os.path.join(download_directory, '_7z')
    if not os.path.isdir(tools_path):
        subprocess.call(
            [exe_path, "x", tool_archive, '-o' + tools_path, '-aoa'])
    return os.path.join(tools_path, '7za.exe')


def extract_old_installer(seven_zip_exe, self_extracting_archive):
    """Extracts the older self-extracting archives.

    seven_zip_exe: The path to 7-zip console executable.
    self_extracting_archive: The executable to extract.
    """
    output_directory = tempfile.mkdtemp()
    subprocess.call(
        [seven_zip_exe, "x", '-o' + output_directory, self_extracting_archive,
         '-i!*.cab'])

    if not os.listdir(output_directory):
        raise ValueError('Failed to extract any cabinet file from ' +
                         self_extracting_archive + '.')

    return output_directory


def fetch_wix(download_directory):
    """Downloads the WiX toolset.

    This is used to extract the runtime packages.
    """
    zip_path = os.path.join(download_directory, '__wix.zip')
    if not os.path.isfile(zip_path):
        wix = 'https://github.com/wixtoolset/wix3/releases/download/' + \
            'wix3111rtm/wix311-binaries.zip'
        download_file(wix, zip_path)

    with zipfile.ZipFile(zip_path, 'r') as zip_file:
        zip_file.extractall(os.path.join(download_directory, '__wix'))

    return os.path.join(download_directory, '__wix')


def extract_burn_bundle(wix_tool_location, bundle_exe):
    """Extracts a burn bundle to a temporary directory.

    bundle_exe: The bundle executable to extract.
    """
    dark_exe = os.path.join(wix_tool_location, 'dark.exe')
    output_directory = tempfile.mkdtemp()
    subprocess.call([dark_exe, "-nologo", "-x", output_directory, bundle_exe])
    return output_directory


def find_cabs(directory):
    base_path = os.path.join(directory, 'AttachedContainer', 'packages')
    for name in os.listdir(base_path):
        if name.endswith('_amd64'):
            yield os.path.join(base_path, name, 'cab1.cab')


def extract_cab(cab_source, destination):
    subprocess.call(['expand.exe', "-F:*", cab_source, destination])


def fetch_all(base_directory, include_old_versions):
    download_directory = os.path.join(base_directory, 'Downloads')

    if not os.path.isdir(download_directory):
        raise SystemError('The download directory does not exist: %s' %
                          download_directory)

    wix_tool_location = fetch_wix(download_directory)
    for ver, installer in fetch_runtimes(download_directory,
                                         include_old_versions):
        output_directory = os.path.join(base_directory, 'vcruntime_' + ver)

        if not os.path.isdir(output_directory):
            os.makedirs(output_directory)

        major_version = int(ver.partition('.')[0])

        if os.listdir(output_directory):
            print('Already have ' + ver)
            continue

        if major_version == 10:
            # The CAB is located located .\.\.\.\vc_red.cab which
            # extract_old_installer() can't cope with.
            print('Cannot extract Visual C++ 2010 runtime. Skipping.')
            continue

        if  major_version < 11:
            # Older versions of the do not use WiX and do not contain a
            # .wixburn data section.
            cab_directory = extract_old_installer(
                fetch_7zip(download_directory), installer)
            for cab in os.listdir(cab_directory):
                extract_cab(os.path.join(cab_directory, cab), output_directory)

            continue

        for cab in find_cabs(extract_burn_bundle(wix_tool_location, installer)):
            extract_cab(cab, output_directory)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Fetch Visual C++ redistributables and extract them.')
    parser.add_argument(
        '--destination',
        help='The destination directory where to download and extract the '
              'files to.',
        default=r'd:\Microsoft\Runtimes')
    parser.add_argument(
        '--include-old-versions',
        action='store_true',
        help='Include older versions of the runtime prior to version 14. The '
             'first version 14 is 2015.')

    arguments = parser.parse_args()
    fetch_all(arguments.destination, arguments.include_old_versions)
