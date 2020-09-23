/*
typedef struct {
  void *pIn;
  void *pOut;
  long w;
  long h;
  long bytes_per_scanline;
  long bytes_per_element;
  long *weights;
} RenderParams
*/

[
    uuid(8350E2A0-B2F7-4a35-9F2D-AEAEEAFBEB06),
    helpstring("Fast Rendering Interface"),
    odl
]
interface IFastRender : IUnknown
{
  void Blur3x3([in] unsigned char *pIn, [in] unsigned char  *pOut, [in] long w, [in] long h, [in] long pitch);
/*
  void SimpleFilter3([in] RenderParams *params);
  void SimpleFilter5([in] RenderParams *params);
  
  void MultiplyNumber([in] RenderParams *params);
  void AddNumber([in] RenderParams *params);
  void SubtractNumber([in] RenderParams *params);
  
  void MultiplyImages([in] RenderParams *params);
  void AddImages([in] RenderParams *params);
  void SubtractImages([in] RenderParams *params);
  
  void DoPremulAlpha([in] RenderParams *params);
  
  void Displace([in] RenderParams *params);   // *pOut is a pointer to a {fixed:fixed} structure array
*/
}