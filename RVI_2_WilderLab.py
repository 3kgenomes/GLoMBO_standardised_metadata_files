#!/usr/bin/env python3
"""Generate a Wilderlab portal workbook from a standardised sampler CSV.

Copies the supplied template package and fills its existing cells. Fails closed
on missing mandatory job/sample fields and duplicate physical kit IDs.
"""
import argparse
import datetime as dt
import re
from pathlib import Path

from glombo_common import (InputError, cell, decimal, patch_workbook,
                           read_json, read_samples, report_csv, samples_for_output,
                           sheet_headers)
from openpyxl import load_workbook


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--samples',type=Path,required=True)
    p.add_argument('--template',type=Path,required=True)
    p.add_argument('--job',type=Path,required=True,help='Job details JSON; explicit data availability')
    p.add_argument('--out',type=Path,required=True)
    p.add_argument('--allow-mock',action='store_true',help='Only for explicit test fixtures')
    args=p.parse_args()
    source=read_samples(args.samples,allow_mock=args.allow_mock)
    rows=samples_for_output(source)
    if not rows:raise InputError('No completed zero-error sampling events')
    config=read_json(args.job)
    if not isinstance(config.get('prepaid',False),bool):
        raise InputError('prepaid must be a JSON boolean')
    mandatory=['ContactName','ContactEmail','Country','AssayPanel','DataAvailability']
    if not config.get('prepaid',False):
        mandatory+=['BillingName','BillingEmail','BillingAddress','ContactPhone']
    missing=[key for key in mandatory if not config.get(key)]
    if missing:raise InputError('Missing Wilderlab job fields: '+', '.join(missing))
    if '@' not in config['ContactEmail']:raise InputError('ContactEmail must contain @')
    if config['DataAvailability'] not in {'Public','Public (after 3-month embargo)','Private'}:
        raise InputError('DataAvailability must be an explicit template option')
    if config.get('ShareResults','No') not in {'Yes','No'}:
        raise InputError('ShareResults must be Yes or No')
    if config.get('ShareResults')=='Yes' and not config.get('ShareResultsAccount'):
        raise InputError('ShareResultsAccount required when ShareResults is Yes')
    wb=load_workbook(args.template,read_only=True,data_only=True)
    for sheet in ('Sample metadata','Extra metadata'):
        if wb[sheet].max_row!=2 or not str(wb[sheet]['A2'].value or '').startswith('Eg:'):
            raise InputError(f'{sheet} is already populated or the blank template layout changed')
    allowed_countries={str(row[0]) for row in wb['Lists'].values if row[0]}
    allowed_panels={str(row[3]) for row in wb['Lists'].values if len(row)>3 and row[3]}
    allowed_environment={str(row[7]) for row in wb['Lists'].values if len(row)>7 and row[7]}
    wb.close()
    if config['Country'] not in allowed_countries:raise InputError('Country not in supplied template list')
    if config['AssayPanel'] not in allowed_panels:raise InputError('AssayPanel not in supplied template list')
    job_book=load_workbook(args.template,read_only=True,data_only=True)
    job_rows={str(row[0]):(index,row[1]) for index,row in enumerate(job_book['Job metadata'].values,1) if row[0]}
    job_book.close()
    sample_fields=sheet_headers(args.template,'Sample metadata',1)
    extra_fields=sheet_headers(args.template,'Extra metadata',1)
    changes={'Job metadata':{},'Sample metadata':{},'Extra metadata':{}}
    for key,(index,current) in job_rows.items():
        if key in config and config[key] is not None:
            changes['Job metadata'][cell(2,index)]=config[key]
        elif current in ('Required','Optional','If known'):
            changes['Job metadata'][cell(2,index)]=None
    sample_map={
        'Kit or UID number':'kit_uid','Site ID':'site_id','Reference':'sample_event_id',
        'Collector':'collector','Date collected':'collection_date',
        'Latitude':'_lat','Longitude':'_lon','Sampling depth':'sampling_depth_m',
        'Environment type':'environment_type'
    }
    review=[]
    for index,r in enumerate(rows,2):
        uid=r['kit_uid'].strip()
        if not re.fullmatch(r'\d{5,6}',uid):
            raise InputError(f"{r['sample_event_id']}: verified Wilderlab Kit/UID must be 5 or 6 digits")
        if r['environment_type'] not in allowed_environment:
            raise InputError(f"{r['sample_event_id']}: unknown Wilderlab environment_type")
        date=(r.get('collection_date') or '').strip()
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}',date):
            raise InputError(
                f"{r['sample_event_id']}: collection_date YYYY-MM-DD required; "
                f"found {date!r} in {args.samples}"
            )
        try:dt.date.fromisoformat(date)
        except ValueError as exc:raise InputError(f"{r['sample_event_id']}: invalid collection_date") from exc
        if r.get('sampling_depth_m',''):
            decimal(r['sampling_depth_m'],0,11000,'sampling_depth_m')
        vals={name:r.get(key,'') for name,key in sample_map.items()}
        vals['Date collected']=date
        vals['Sample amount']=f"{r['_volume']:g} mL" if r['_volume'] is not None else ''
        vals['Extra notes']=f"Sampler event {r['sample_event_id']}; filter ID {r['filter_id']}; position {r['filter_position']}; source {r.get('source_log_file','')}"
        for name,colnum in sample_fields.items():
            changes['Sample metadata'][cell(colnum,index)]=vals.get(name,'')
        # The provided Extra metadata sheet enables a verified reverse join.
        changes['Extra metadata'][cell(extra_fields['Kit or UID number'],index)]=uid
        changes['Extra metadata'][cell(extra_fields['Metadata name'],index)]='Sampler event ID'
        changes['Extra metadata'][cell(extra_fields['Metadata value'],index)]=r['sample_event_id']
        review.append({'sample_event_id':r['sample_event_id'],'kit_uid':uid,
                       'status':'READY' if r['_lat'] is not None else 'NO_MAP_COORDINATES',
                       'note':'' if r['_lat'] is not None else 'Date supplied; no WGS84 position for explore map'})
    patch_workbook(args.template,args.out,changes)
    report_csv(args.out.with_suffix('.review.csv'),['sample_event_id','kit_uid','status','note'],review)
    print(f'Wilderlab workbook: {len(rows)} completed samples; {len(source)-len(rows)} attempts excluded')


if __name__=='__main__':
    try:main()
    except InputError as exc:raise SystemExit(f'Input error: {exc}')
